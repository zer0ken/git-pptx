'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');

const deckMod = require('./deck');
const { diffDirs, listFiles, isDerived } = require('./diff');
const { normalizePartNames } = require('./normalize');
const { renderSlides } = require('./render');

// ANSI styling that degrades to plain text when the stream is not a TTY (e.g.
// when piped or redirected). Progress goes to stderr (dim), results to stdout
// (bold) so the key outcome stands out from the running commentary.
function makeStyle(stream) {
  const on = () => !!stream.isTTY;
  return {
    bold: (s) => (on() ? '\x1b[1m' + s + '\x1b[0m' : s),
    dim: (s) => (on() ? '\x1b[2m' + s + '\x1b[0m' : s),
  };
}
const out = makeStyle(process.stdout);
const err = makeStyle(process.stderr);

function joinList(items, max) {
  if (items.length <= max) return items.join(', ');
  return items.slice(0, max).join(', ') + ', ... (' + items.length + ' total)';
}

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

// Unpack into a scratch directory and give the parts their canonical names, so
// that the same deck always lands on the same filenames no matter which tool
// wrote it. --no-normalize keeps whatever names the deck came with.
function unpackFor(deck, dir, flags) {
  deckMod.unzipToDir(deck, dir);
  if (!flags['no-normalize']) normalizePartNames(dir);
}

// Flags that carry a value, written either --format=png or --format png. Without
// this list the second form leaves "png" in the positional arguments, where
// decomp would take it for the output directory.
const VALUE_FLAGS = new Set(['format', 'renderer']);

function parseArgs(args) {
  const flags = {};
  const positional = [];
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (!a.startsWith('--')) {
      positional.push(a);
      continue;
    }
    const eq = a.indexOf('=');
    if (eq !== -1) {
      flags[a.slice(2, eq)] = a.slice(eq + 1);
      continue;
    }
    const name = a.slice(2);
    if (VALUE_FLAGS.has(name) && i + 1 < args.length && !args[i + 1].startsWith('--')) {
      flags[name] = args[++i];
    } else {
      flags[name] = true;
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
    throw new Error('usage: git-pptx decomp <deck.pptx> [dir] [--no-preview] [--format jpg|png] [--no-normalize]');
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
  const log = (m) => process.stderr.write(err.dim(m) + '\n');
  let rendered = 0;
  try {
    log('unzipping ' + path.basename(deck) + '...');
    unpackFor(deck, tmp, flags);
    log(existing ? 'comparing with ' + dir + '...' : 'splitting files...');
    // Everything that differs is written, derived parts included, so the
    // directory always holds a package comp can rebuild. Only `reported` is
    // what the user is told changed and what previews are re-rendered for.
    const changed = existing ? diffDirs(tmp, pptxDir) : listFiles(tmp);
    if (changed.length) {
      log('copying ' + changed.length + ' changed file(s)...');
      copyChanged(changed, tmp, pptxDir);
    }
    const reported = changed.filter((rel) => !isDerived(rel));

    if (!flags['no-preview']) {
      const idx = changedSlideIndices(reported);
      if (idx.length) {
        // idx holds slide PART numbers (the N in slideN.xml), which is only the
        // part's own creation order, not its position in this deck's slide show.
        // A deck split out of a bigger one keeps its original part numbers, so
        // slide3.xml can be the 1st slide of a 6-slide file. Resolve each part
        // number to its real 1-based position before asking PowerPoint or
        // LibreOffice to export it: otherwise the wrong slide is rendered, and
        // an out-of-range position makes the export fail outright.
        const posMap = deckMod.slidePartToPosition(pptxDir);
        const pairs = idx.map((n) => ({ label: n, position: posMap[n] })).filter((p) => p.position);
        const unresolved = idx.filter((n) => !posMap[n]);
        if (unresolved.length) {
          console.warn('warning: could not resolve slide position for: ' + unresolved.join(', ') + ' (skipped preview)');
        }
        if (pairs.length) {
          // Render from a temp copy so an open editor is never touched.
          const renderCopy = path.join(tmp, 'deck-render.pptx');
          fs.copyFileSync(deck, renderCopy);
          log('rendering ' + pairs.length + ' preview(s)...');
          await renderSlides(renderCopy, previewsDir, pairs, format, flags.renderer, log);
          rendered = pairs.length;
        }
      }
    }

    if (reported.length === 0) {
      console.log(out.bold('no changes'));
    } else {
      // Summarize: name the changed slides, collapse everything else into a
      // count so media/metadata churn stays out of the way.
      const posMap = deckMod.slidePartToPosition(pptxDir);
      const chgSlides = changedSlideIndices(reported).map((n) => posMap[n] || n);
      const otherCount = reported.length - chgSlides.length;
      console.log(out.bold('updated ' + reported.length + ' file(s)'));
      if (chgSlides.length) {
        console.log(out.dim('  ' + chgSlides.length + ' slide(s): ' + joinList(chgSlides.map((p) => 'slide ' + p), 20)));
      }
      if (otherCount) {
        console.log(out.dim('  + ' + otherCount + ' other file(s)'));
      }
    }
    if (rendered) {
      console.log(out.bold('rendered ' + rendered + ' preview(s)'));
    }
    console.log(out.bold('git-pptx dir:') + ' ' + dir);
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
  const outPath = path.resolve(positional[1]);
  const pptxDir = path.join(dir, PPTX_SUB);
  if (!fs.existsSync(pptxDir)) throw new Error('not a git-pptx dir: ' + dir);
  deckMod.zipDirToFile(pptxDir, outPath);
  console.log(out.bold('composed') + ' ' + outPath);
}

async function cmdDiff(args) {
  const { flags, positional } = parseArgs(args);
  if (positional.length < 2) {
    throw new Error('usage: git-pptx diff <deck.pptx> <dir> [--no-normalize]');
  }
  const deck = path.resolve(positional[0]);
  const dir = path.resolve(positional[1]);
  if (!fs.existsSync(deck)) throw new Error('deck not found: ' + deck);
  const pptxDir = path.join(dir, PPTX_SUB);
  fs.mkdirSync(pptxDir, { recursive: true });
  const tmp = tmpdir();
  try {
    unpackFor(deck, tmp, flags);
    const changed = diffDirs(tmp, pptxDir).filter((rel) => !isDerived(rel));
    if (changed.length === 0) {
      console.log(out.bold('in sync'));
      return;
    }
    // Summarize like decomp: name the changed slides, collapse the rest into a
    // count so media/metadata churn stays out of the way.
    const posMap = deckMod.slidePartToPosition(pptxDir);
    const chgSlides = changedSlideIndices(changed).map((n) => posMap[n] || n);
    const otherCount = changed.length - chgSlides.length;
    console.log(out.bold('diff: ' + changed.length + ' file(s)'));
    if (chgSlides.length) {
      console.log(out.dim('  ' + chgSlides.length + ' slide(s): ' + joinList(chgSlides.map((p) => 'slide ' + p), 20)));
    }
    if (otherCount) {
      console.log(out.dim('  + ' + otherCount + ' other file(s)'));
    }
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

Numbered parts (slide7.xml, image240.png, theme4.xml) are renamed to the dense
1..N sequence PowerPoint itself uses, so a deck written by a script or split out
of a bigger deck does not have every part renamed - and every file in the
directory rewritten - the first time someone saves it in PowerPoint. Pass
--no-normalize to keep the names the deck came with.

Usage:
  git-pptx decomp <deck.pptx> [dir] [--no-preview] [--format jpg|png] [--no-normalize]
      split deck.pptx into <deck>.git-pptx/ (or dir). [--renderer auto|powerpoint|libreoffice]. If the pptx/ subdir
      already holds a previous split, only changed files are updated and
      previews are rendered only for changed slides. Default format is jpg.
  git-pptx comp <dir> <deck.pptx>
      rebuild deck.pptx from the pptx/ subdir of a git-pptx dir.
  git-pptx diff <deck.pptx> <dir> [--no-normalize]
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
