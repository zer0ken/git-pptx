'use strict';

const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

// pptx is a zip of an XML directory. "Splitting" is a lossless unzip and
// "merging" is a lossless rezip of that directory. No reconstruction of
// relationship IDs or [Content_Types].xml happens here, so no data is lost.

function unzipToDir(pptxPath, dir) {
  const zip = new AdmZip(pptxPath);
  for (const e of zip.getEntries()) {
    if (e.isDirectory) continue;
    const target = path.join(dir, e.entryName);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, e.getData());
  }
}

// The numeric suffix in an OOXML slide part filename (ppt/slides/slideN.xml)
// is just the order the part was originally created in; it is NOT the slide's
// position in this deck's slide show. Decks assembled or split from a larger
// deck (this is the normal case for the split section files this project
// keeps under slides/) keep their original part numbers, so slide3.xml can
// easily be the 1st slide of a 6-slide file. Anything that asks PowerPoint or
// LibreOffice to export "slide N" needs the real 1-based position, which only
// ppt/presentation.xml's <p:sldIdLst> (via ppt/_rels/presentation.xml.rels)
// actually records.
function getSlidePartOrder(pptxDir) {
  const presXml = fs.readFileSync(path.join(pptxDir, 'ppt', 'presentation.xml'), 'utf8');
  const relsPath = path.join(pptxDir, 'ppt', '_rels', 'presentation.xml.rels');
  const relsXml = fs.readFileSync(relsPath, 'utf8');

  const relTargets = {};
  for (const tag of relsXml.match(/<Relationship\b[^>]*\/>/g) || []) {
    const id = (tag.match(/\bId="([^"]+)"/) || [])[1];
    const target = (tag.match(/\bTarget="([^"]+)"/) || [])[1];
    if (id && target) relTargets[id] = target;
  }

  const order = [];
  for (const tag of presXml.match(/<p:sldId\b[^>]*\/>/g) || []) {
    const rid = (tag.match(/r:id="([^"]+)"/) || [])[1];
    const target = rid && relTargets[rid];
    if (target) order.push(path.basename(target));
  }
  return order; // order[0] is the part filename at position 1, etc.
}

// Maps each slide part's filename number (the N in slideN.xml) to its real
// 1-based position in the deck's slide show order.
function slidePartToPosition(pptxDir) {
  const map = {};
  getSlidePartOrder(pptxDir).forEach((name, i) => {
    const m = name.match(/^slide(\d+)\.xml$/);
    if (m) map[+m[1]] = i + 1;
  });
  return map;
}

function zipDirToFile(dir, pptxPath) {
  const zip = new AdmZip();
  const FIXED = new Date(1980, 0, 1, 0, 0, 0);
  (function walk(d, rel) {
    for (const entry of fs.readdirSync(d, { withFileTypes: true })) {
      const full = path.join(d, entry.name);
      const name = rel ? rel + '/' + entry.name : entry.name;
      if (entry.isDirectory()) {
        if (entry.name === 'preview') continue; // derived, not part of the pptx
        if (rel === '' && entry.name === '.git') continue; // VCS metadata
        walk(full, name);
      } else {
        if (rel === '' && entry.name.startsWith('.')) continue; // top-level VCS dotfiles
        zip.addFile(name, fs.readFileSync(full), '', 0o644, FIXED);
      }
    }
  })(dir, '');
  zip.writeZip(pptxPath);
}

module.exports = { unzipToDir, zipDirToFile, getSlidePartOrder, slidePartToPosition };
