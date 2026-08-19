'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');

const deckMod = require('./deck');
const { diffDirs, listFiles } = require('./diff');
const { renderSlides } = require('./render');

// git-pptx directory layout for input "a.pptx":
//   a.git-pptx/
//     previews/   1.jpg, 2.jpg, ...   (per-slide previews, derived)
//     pptx/       the pptx in unpacked form
const PPTX_SUB = 'pptx';
const PREVIEWS_SUB = 'previews';

// Derive the git-pptx directory name from a deck filename: a.pptx -> a.git-pptx
function defaultDir(deck) {
  const base = path.basename(deck, path.extname(deck));
  return path.join(path.dirname(deck), base + '.git-pptx');
}

function tmpdir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'gpptx-'));
}

function slideLabel(rel) {
  const m = rel.match(/^ppt\/slides\/slide(\d+)\.xml$/);
  if (m) return 'slide ' + m[1];
  return rel;
}

function parseArgs(args) {
  const flags = {};
  const positional = [];
  for (const a of args) {
    if (a.startsWith('--')) {
      const eq = a.indexOf('=');
      if (eq === -1) flags[a.slice(2)] = true;
      else flags[a.slice(2, eq)] = a.slice(eq + 1);
    } else {
      positional.push(a);
    }
  }
  return { flags, positional };
}

function changedSlideIndices(changed) {
  return [...new Set(
    changed
      .map((rel) => rel.match(/^ppt\/slides\/slide(\d+)\.xml$/))
      .filter(Boolean)
      .map((m) => +m[1])
  )].sort((a, b) => a - b);
}

function copyChanged(changed, fromDir, toDir) {
  for (const rel of changed) {
    const src = path.join(fromDir, rel);
    const dst = path.join(toDir, rel);
    if (fs.existsSync(src)) {
      fs.mkdirSync(path.dirname(dst), { recursive: true });
      fs.copyFileSync(src, dst);
    } else {
      fs.rmSync(dst, { force: true });
    }
  }
}

// Decompose a pptx into a git-pptx directory. The output directory defaults to
// the deck filename (a.pptx -> a.pptx/). When it already holds a previous
// split, only files that actually changed are updated and previews are
// rendered only for the changed slides.
async function cmdDecomp(args) {
  const { flags, positional } = parseArgs(args);
  if (positional.length < 1) {
    throw new Error('usage: git-pptx decomp <deck.pptx> [dir] [--no-preview] [--format jpg|png]');
  }
  const deck = path.resolve(positional[0]);
  if (!fs.existsSync(deck)) throw new Error('deck not found: ' + deck);
  const dir = positional[1]
    ? path.resolve(positional[1])
    : defaultDir(deck);
  const format = flags.format === 'png' ? 'png' : 'jpg';
  const pptxDir = path.join(dir, PPTX_SUB);
  const previewsDir = path.join(dir, PREVIEWS_SUB);

  const existing = fs.existsSync(pptxDir) && fs.readdirSync(pptxDir).length > 0;
  const tmp = tmpdir();
  try {
    deckMod.unzipToDir(deck, tmp);
    const changed = existing ? diffDirs(tmp, pptxDir) : listFiles(tmp);
    copyChanged(changed, tmp, pptxDir);

    if (!flags['no-preview']) {
      const idx = changedSlideIndices(changed);
      if (idx.length) {
        renderSlides(deck, previewsDir, idx, format);
        console.log('rendered previews for ' + idx.length + ' slide(s)');
      }
    }

    if (changed.length === 0) {
      console.log('no changes');
    } else {
      console.log('updated ' + changed.length + ' file(s): ' + changed.map(slideLabel).join(', '));
    }
    console.log('git-pptx dir: ' + dir);
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

async function cmdComp(args) {
  const { positional } = parseArgs(args);
  if (positional.length < 2) {
    throw new Error('usage: git-pptx comp <dir> <deck.pptx>');
  }
  const dir = path.resolve(positional[0]);
  const out = path.resolve(positional[1]);
  const pptxDir = path.join(dir, PPTX_SUB);
  if (!fs.existsSync(pptxDir)) throw new Error('not a git-pptx dir: ' + dir);
  deckMod.zipDirToFile(pptxDir, out);
  console.log('composed ' + out);
}

async function cmdDiff(args) {
  const { positional } = parseArgs(args);
  if (positional.length < 2) {
    throw new Error('usage: git-pptx diff <deck.pptx> <dir>');
  }
  const deck = path.resolve(positional[0]);
  const dir = path.resolve(positional[1]);
  if (!fs.existsSync(deck)) throw new Error('deck not found: ' + deck);
  const pptxDir = path.join(dir, PPTX_SUB);
  fs.mkdirSync(pptxDir, { recursive: true });
  const tmp = tmpdir();
  try {
    deckMod.unzipToDir(deck, tmp);
    const changed = diffDirs(tmp, pptxDir);
    if (changed.length === 0) {
      console.log('in sync');
      return;
    }
    for (const rel of changed) console.log('* ' + slideLabel(rel));
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

const HELP = `git-pptx - decompose and compose PowerPoint (pptx) as a git-friendly directory

For input a.pptx the git-pptx directory is created as a.git-pptx/ containing:

  a.git-pptx/
    previews/   1.jpg, 2.jpg, ...   per-slide previews (derived)
    pptx/       the pptx in unpacked form

The directory name is the deck name with the extension replaced by .git-pptx,
so it can sit next to the .pptx file without colliding.

Usage:
  git-pptx decomp <deck.pptx> [dir] [--no-preview] [--format jpg|png]
      split deck.pptx into <deck>.git-pptx/ (or dir). If the pptx/ subdir
      already holds a previous split, only changed files are updated and
      previews are rendered only for changed slides. Default format is jpg.
  git-pptx comp <dir> <deck.pptx>
      rebuild deck.pptx from the pptx/ subdir of a git-pptx dir.
  git-pptx diff <deck.pptx> <dir>
      show which slides differ without writing anything.

There is no git or GitHub coupling; publish the directory with ordinary git:

  git add a.git-pptx && git commit && git push
`;

async function run(args) {
  const cmd = args[0] || 'help';
  const rest = args.slice(1);
  switch (cmd) {
    case 'decomp': return cmdDecomp(rest);
    case 'comp': return cmdComp(rest);
    case 'diff': return cmdDiff(rest);
    case 'help':
    case '--help':
    case '-h': console.log(HELP); return;
    default: throw new Error('unknown command: ' + cmd + '\n' + HELP);
  }
}

module.exports = { run };
