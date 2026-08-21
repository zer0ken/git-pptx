'use strict';

// Self-check for preview freshness. Runs the decomp/comp loop over the generated
// fixture with the renderer stubbed out, so it needs neither PowerPoint nor
// LibreOffice:
//   node scripts/check-previews.js
//
// The case that matters: comp writes a pptx that matches the work dir exactly,
// so the pptx-against-work-dir comparison sees no change on the decomp that
// follows, while the previews still show the state before the edit.

const fs = require('fs');
const os = require('os');
const path = require('path');

// Stub the renderer before commands.js is loaded: it binds renderSlides at
// require time, so the stub has to be in place first.
const rendered = [];
require('../lib/render').renderSlides = async (deckPath, outDir, pairs, format) => {
  fs.mkdirSync(outDir, { recursive: true });
  for (const { label } of pairs) {
    fs.writeFileSync(path.join(outDir, label + '.' + format), 'image of slide ' + label);
    rendered.push(label);
  }
};

const commands = require('../lib/commands');
const { make } = require('./make-fixture');

function fail(msg) {
  console.error('FAIL: ' + msg);
  process.exit(1);
}

function same(actual, expected, what) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) fail(what + ': expected ' + e + ', got ' + a);
}

// decomp writes its progress to stderr and its summary to stdout; both are
// noise here, so only the list of slides it rendered comes back.
async function decomp(args) {
  const log = console.log;
  const warn = console.warn;
  const write = process.stderr.write.bind(process.stderr);
  console.log = () => {};
  console.warn = () => {};
  process.stderr.write = () => true;
  rendered.length = 0;
  try {
    await commands.run(['decomp'].concat(args));
  } finally {
    console.log = log;
    console.warn = warn;
    process.stderr.write = write;
  }
  return rendered.slice();
}

async function comp(dir, deck) {
  const log = console.log;
  console.log = () => {};
  try {
    await commands.run(['comp', dir, deck]);
  } finally {
    console.log = log;
  }
}

async function main() {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'gpptx-check-'));
  const deck = path.join(tmp, 'deck.pptx');
  const dir = path.join(tmp, 'deck.git-pptx');
  const slide1 = path.join(dir, 'pptx', 'ppt', 'slides', 'slide1.xml');
  try {
    make(deck, ['Slide One', 'Slide Two', 'Slide Three']);

    same(await decomp([deck, dir]), [1, 2, 3], 'first decomp renders every slide');
    same(await decomp([deck, dir]), [], 'unchanged deck renders nothing');

    // The reported issue: edit a part, comp, decomp. The pptx and the work dir
    // agree at that point, so only the recorded fingerprint can tell that the
    // preview predates the edit.
    fs.writeFileSync(slide1, fs.readFileSync(slide1, 'utf8').replace('Slide One', 'Edited One'));
    await comp(dir, deck);
    same(await decomp([deck, dir]), [1], 'edited slide is rendered again after comp');
    same(await decomp([deck, dir]), [], 'and stays fresh afterwards');

    // Reserialization noise is not a content change, so it must not re-render.
    fs.writeFileSync(slide1, fs.readFileSync(slide1, 'utf8').replace(/<a:t>/g, '<a:t >'));
    await comp(dir, deck);
    same(await decomp([deck, dir]), [], 'reserialization noise renders nothing');

    // Switching format renders everything again and leaves no image behind in
    // the format that is no longer asked for.
    same(await decomp([deck, dir, '--format', 'png']), [1, 2, 3], 'format switch renders every slide');
    const left = fs.readdirSync(path.join(dir, 'previews')).sort();
    same(left, ['1.png', '2.png', '3.png', 'index.json'], 'previews of the old format are dropped');

    console.log('ok: previews track the content they were rendered from across comp/decomp');
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

main().catch((e) => fail(e.stack || e.message));
