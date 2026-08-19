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

module.exports = { unzipToDir, zipDirToFile };
