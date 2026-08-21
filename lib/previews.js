'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const { canonicalize } = require('./canonical');

// Whether a preview is current is a question about the content it was rendered
// from, not about how the work dir compares to the pptx: right after comp the
// unpacked pptx and the work dir are identical by construction, so a comparison
// between them reports no change even when the previews predate the edit that
// produced the pptx. decomp therefore records the fingerprint each preview was
// rendered at and re-renders whenever the slide no longer matches it.
const MANIFEST = 'index.json';

// Only the files decomp itself writes are ever removed from previews/.
const PREVIEW_FILE = /^\d+\.(jpg|png)$/;

// The same notion of content the change report uses (lib/diff.js), so
// reserialization noise and refreshed field caches do not force a re-render.
function fingerprint(file) {
  const raw = fs.readFileSync(file);
  let text = null;
  try {
    text = canonicalize(raw.toString('utf8'));
  } catch (e) {
    text = null;
  }
  return crypto.createHash('sha256').update(text === null ? raw : text).digest('hex').slice(0, 16);
}

// Slide parts as they sit in the unpacked package. `label` is the part's own
// number (the N in slideN.xml), which is what the preview file is named after;
// it is unrelated to the slide's position in the slide show.
function slideParts(pptxDir) {
  const dir = path.join(pptxDir, 'ppt', 'slides');
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir)
    .map((name) => name.match(/^slide(\d+)\.xml$/))
    .filter(Boolean)
    .map((m) => ({ label: +m[1], file: path.join(dir, m[0]) }))
    .sort((a, b) => a.label - b.label);
}

function readManifest(previewsDir) {
  try {
    const m = JSON.parse(fs.readFileSync(path.join(previewsDir, MANIFEST), 'utf8'));
    return { format: m.format, slides: m.slides || {} };
  } catch (e) {
    return { format: null, slides: {} };
  }
}

// Which previews need rendering: the slide's fingerprint differs from the one
// its preview was rendered at, the image is missing, or the images on disk are
// in a different format than the one asked for.
function planPreviews(pptxDir, previewsDir, format) {
  const manifest = readManifest(previewsDir);
  const rendered = manifest.format === format ? manifest.slides : {};
  const wanted = {};
  const stale = [];
  for (const { label, file } of slideParts(pptxDir)) {
    const fp = fingerprint(file);
    wanted[label] = fp;
    const img = path.join(previewsDir, label + '.' + format);
    if (rendered[String(label)] !== fp || !fs.existsSync(img)) stale.push(label);
  }
  return { wanted, stale };
}

// Records what the previews on disk were rendered from and drops the ones that
// no longer belong to a slide (a deleted slide, or an image left behind by an
// earlier --format). A fingerprint is recorded only for an image that is
// actually there, so a render that failed part way through is retried instead
// of being remembered as current.
function commitPreviews(previewsDir, format, wanted) {
  // No slide parts means nothing is known about what the previews belong to,
  // so nothing is dropped and nothing is recorded.
  if (!fs.existsSync(previewsDir) || Object.keys(wanted).length === 0) return { removed: [] };
  const keep = new Set(Object.keys(wanted).map((label) => label + '.' + format));
  const removed = [];
  for (const name of fs.readdirSync(previewsDir)) {
    if (!PREVIEW_FILE.test(name) || keep.has(name)) continue;
    fs.rmSync(path.join(previewsDir, name), { force: true });
    removed.push(name);
  }
  const slides = {};
  for (const [label, fp] of Object.entries(wanted)) {
    if (fs.existsSync(path.join(previewsDir, label + '.' + format))) slides[label] = fp;
  }
  fs.writeFileSync(
    path.join(previewsDir, MANIFEST),
    JSON.stringify({ format, slides }, null, 2) + '\n'
  );
  return { removed };
}

module.exports = { planPreviews, commitPreviews };
