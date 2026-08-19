'use strict';

// Self-check for canonical part naming. Builds an unpacked package whose parts
// carry the sparse numbers a deck split out of a bigger one ends up with,
// normalizes it, and asserts that the parts land on the dense names PowerPoint
// would have used and that nothing dangles. No PowerPoint or LibreOffice needed:
//   node scripts/check-normalize.js

const fs = require('fs');
const os = require('os');
const path = require('path');
const { normalizePartNames } = require('../lib/normalize');

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P = 'http://schemas.openxmlformats.org/package/2006/relationships';

function rels(list) {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="' + P + '">'
    + list.map((r) => '<Relationship Id="' + r[0] + '" Type="' + r[1] + '" Target="' + r[2] + '"/>').join('')
    + '</Relationships>';
}

// A one-slide deck cut out of a 74-slide one: the slide is still slide74.xml,
// its layout slideLayout14.xml, its pictures image240/241/6.
const FILES = {
  '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Default Extension="png" ContentType="image/png"/>'
    + '<Default Extension="jpg" ContentType="image/jpeg"/>'
    + '<Override PartName="/ppt/presentation.xml" ContentType="pres"/>'
    + '<Override PartName="/ppt/slides/slide74.xml" ContentType="slide"/>'
    + '<Override PartName="/ppt/slideLayouts/slideLayout14.xml" ContentType="layout"/>'
    + '<Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="master"/>'
    + '<Override PartName="/ppt/theme/theme4.xml" ContentType="theme"/>'
    + '</Types>',
  '_rels/.rels': rels([['rId1', R + '/officeDocument', 'ppt/presentation.xml']]),
  'ppt/presentation.xml': '<p:presentation/>',
  'ppt/_rels/presentation.xml.rels': rels([
    ['rId5', R + '/slideMaster', 'slideMasters/slideMaster2.xml'],
    ['rId80', R + '/slide', 'slides/slide74.xml'],
    ['rId105', R + '/theme', 'theme/theme4.xml'],
  ]),
  'ppt/slides/slide74.xml': '<p:sld/>',
  'ppt/slides/_rels/slide74.xml.rels': rels([
    ['rId1', R + '/slideLayout', '../slideLayouts/slideLayout14.xml'],
    ['rId2', R + '/image', '../media/image240.jpg'],
    ['rId3', R + '/image', '../media/image241.png'],
    ['rId4', R + '/image', '../media/image6.png'],
  ]),
  'ppt/slideLayouts/slideLayout14.xml': '<p:sldLayout/>',
  'ppt/slideLayouts/_rels/slideLayout14.xml.rels': rels([
    ['rId1', R + '/slideMaster', '../slideMasters/slideMaster2.xml'],
  ]),
  'ppt/slideMasters/slideMaster2.xml': '<p:sldMaster/>',
  'ppt/slideMasters/_rels/slideMaster2.xml.rels': rels([
    ['rId2', R + '/slideLayout', '../slideLayouts/slideLayout14.xml'],
    ['rId3', R + '/theme', '../theme/theme4.xml'],
  ]),
  'ppt/theme/theme4.xml': '<a:theme/>',
  'ppt/media/image240.jpg': 'jpg-bytes',
  'ppt/media/image241.png': 'png-bytes-a',
  'ppt/media/image6.png': 'png-bytes-b',
};

// Extensions share a prefix's counter, so image240.jpg/image241.png/image6.png
// become image1.jpg/image2.png/image3.png in the order the slide references them.
const EXPECTED = {
  'ppt/slides/slide1.xml': '<p:sld/>',
  'ppt/slideLayouts/slideLayout1.xml': '<p:sldLayout/>',
  'ppt/slideMasters/slideMaster1.xml': '<p:sldMaster/>',
  'ppt/theme/theme1.xml': '<a:theme/>',
  'ppt/media/image1.jpg': 'jpg-bytes',
  'ppt/media/image2.png': 'png-bytes-a',
  'ppt/media/image3.png': 'png-bytes-b',
};

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'gpptx-check-'));
for (const [rel, body] of Object.entries(FILES)) {
  const p = path.join(dir, rel);
  fs.mkdirSync(path.dirname(p), { recursive: true });
  fs.writeFileSync(p, body);
}

const failures = [];
const renamed = normalizePartNames(dir);
if (renamed !== 7) failures.push('expected 7 parts renamed, got ' + renamed);

for (const [rel, body] of Object.entries(EXPECTED)) {
  const p = path.join(dir, rel);
  if (!fs.existsSync(p)) failures.push('missing ' + rel);
  else if (fs.readFileSync(p, 'utf8') !== body) failures.push(rel + ' holds the wrong part');
}
for (const rel of Object.keys(FILES)) {
  if (EXPECTED[rel] === undefined && rel.startsWith('ppt/') && !rel.includes('/_rels/')
      && rel !== 'ppt/presentation.xml' && fs.existsSync(path.join(dir, rel))) {
    failures.push(rel + ' should have been renamed away');
  }
}

// Every relationship target and every content-type override has to resolve.
const list = [];
(function walk(d, rel) {
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const name = rel ? rel + '/' + e.name : e.name;
    if (e.isDirectory()) walk(path.join(d, e.name), name);
    else list.push(name);
  }
})(dir, '');
const present = new Set(list);
for (const rel of list) {
  if (!rel.endsWith('.rels')) continue;
  const ownerDir = path.posix.dirname(path.posix.dirname(rel));
  const xml = fs.readFileSync(path.join(dir, rel), 'utf8');
  for (const [, target] of xml.matchAll(/Target="([^"]+)"/g)) {
    const base = ownerDir === '.' ? '' : ownerDir + '/';
    const resolved = target.startsWith('/') ? target.slice(1) : path.posix.normalize(base + target);
    if (!present.has(resolved)) failures.push(rel + ' -> ' + target + ' does not resolve');
  }
}
const ct = fs.readFileSync(path.join(dir, '[Content_Types].xml'), 'utf8');
for (const [, partName] of ct.matchAll(/PartName="([^"]+)"/g)) {
  if (!present.has(partName.replace(/^\//, ''))) failures.push('content type for missing ' + partName);
}

// A package that is already canonical must come out untouched.
const again = normalizePartNames(dir);
if (again !== 0) failures.push('normalizing twice renamed ' + again + ' more parts');

fs.rmSync(dir, { recursive: true, force: true });
if (failures.length) {
  console.error('FAIL');
  for (const f of failures) console.error('  ' + f);
  process.exit(1);
}
console.log('ok: sparse part names normalized to ' + Object.keys(EXPECTED).length
  + ' canonical names, references intact, second run a no-op');
