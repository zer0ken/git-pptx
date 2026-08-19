'use strict';

const fs = require('fs');
const path = require('path');
const { canonicalize } = require('./canonical');

// Parts PowerPoint regenerates on every save that carry no slide content:
// document properties it recomputes (word/paragraph counts, the list of slide
// titles), the preview thumbnail, the re-embedded font blobs, and the
// undo/change-tracking parts.
//
// These are still written to the directory. Leaving them out is not an option:
// [Content_Types].xml, ppt/_rels/presentation.xml.rels and ppt/presentation.xml
// all point at them, and those parts ARE written, so skipping the targets
// leaves dangling relationships and PowerPoint refuses to open the composed
// deck. They are only left out of the *reported* change set, so a save that
// touched nothing else does not show up as a diff.
const DERIVED = new Set(['docProps/core.xml', 'docProps/app.xml', 'docProps/thumbnail.jpeg']);
const DERIVED_PREFIX = ['ppt/fonts/', 'ppt/changesInfos/'];

function isDerived(rel) {
  if (DERIVED.has(rel)) return true;
  return DERIVED_PREFIX.some((p) => rel.startsWith(p));
}

function listFiles(dir) {
  const out = [];
  (function walk(d, rel) {
    for (const entry of fs.readdirSync(d, { withFileTypes: true })) {
      const full = path.join(d, entry.name);
      const name = rel ? rel + '/' + entry.name : entry.name;
      if (entry.isDirectory()) walk(full, name);
      else out.push(name);
    }
  })(dir, '');
  return out.sort();
}

function isXml(rel) {
  return /\.(xml|rels)$/.test(rel) || path.basename(rel) === '[Content_Types].xml';
}

function contentEqual(a, b, xml) {
  if (xml) {
    try {
      return canonicalize(fs.readFileSync(a, 'utf8')) === canonicalize(fs.readFileSync(b, 'utf8'));
    } catch (e) {
      return false;
    }
  }
  return fs.readFileSync(a).equals(fs.readFileSync(b));
}

// Returns the list of relative paths whose content differs between aDir and
// bDir, including derived parts (callers filter those out of what they report
// with isDerived, but must still write them). Handles added and removed files.
function diffDirs(aDir, bDir) {
  const a = listFiles(aDir);
  const b = listFiles(bDir);
  const changed = [];
  for (const rel of new Set([...a, ...b])) {
    const pa = path.join(aDir, rel);
    const pb = path.join(bDir, rel);
    const ea = fs.existsSync(pa);
    const eb = fs.existsSync(pb);
    if (!ea || !eb) {
      changed.push(rel);
      continue;
    }
    if (!contentEqual(pa, pb, isXml(rel))) changed.push(rel);
  }
  return changed.sort();
}

module.exports = { diffDirs, listFiles, isXml, isDerived };
