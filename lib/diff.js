'use strict';

const fs = require('fs');
const path = require('path');
const { canonicalize } = require('./canonical');

// Volatile files that change on every PowerPoint save but carry no slide
// content. They are excluded from change detection, so they do not produce
// spurious diffs. They are still kept in the directory for comp fidelity.
const VOLATILE = new Set(['docProps/core.xml']);
const VOLATILE_PREFIX = ['ppt/fonts/', 'ppt/changesInfos/'];

function isVolatile(rel) {
  if (VOLATILE.has(rel)) return true;
  return VOLATILE_PREFIX.some((p) => rel.startsWith(p));
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
// bDir (ignoring volatile files). Handles added and removed files.
function diffDirs(aDir, bDir) {
  const a = listFiles(aDir);
  const b = listFiles(bDir);
  const changed = [];
  for (const rel of new Set([...a, ...b])) {
    if (isVolatile(rel)) continue;
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
  return changed;
}

module.exports = { diffDirs, listFiles, isXml };
