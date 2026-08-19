'use strict';

// Simulates two PowerPoint edits on the generated fixture:
//  - slides 2 and 3: real text content change
//  - slide 1: only reserialization noise (attribute order, empty-element
//    notation, entity encoding) with identical semantic content
// Rewrites deck.pptx in place.

const fs = require('fs');
const os = require('os');
const path = require('path');
const AdmZip = require('adm-zip');

const deckPath = process.argv[2] || 'deck.pptx';
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pit-mut-'));

const zip = new AdmZip(deckPath);
const entries = {};
for (const e of zip.getEntries()) entries[e.entryName] = e.getData().toString('utf8');

function readSlide(i) { return entries[`ppt/slides/slide${i}.xml`]; }

// Slide 2: real text change
entries['ppt/slides/slide2.xml'] = readSlide(2).replace('Slide Two', 'Slide Two EDITED');

// Slide 3: real text change
entries['ppt/slides/slide3.xml'] = readSlide(3).replace('Slide Three', 'Slide Three EDITED');

// Slide 1: noise only, same semantic content
const s1 = readSlide(1)
  .replace('<p:cNvPr id="1" name=""/>', '<p:cNvPr name="" id="1"></p:cNvPr>')
  .replace('<p:grpSpPr/>', '<p:grpSpPr></p:grpSpPr>')
  .replace('<p:nvPr/>', '<p:nvPr></p:nvPr>')
  .replace('<p:cNvSpPr/>', '<p:cNvSpPr></p:cNvSpPr>')
  .replace('<p:spPr/>', '<p:spPr></p:spPr>');
entries['ppt/slides/slide1.xml'] = s1;

const out = new AdmZip();
for (const [name, content] of Object.entries(entries)) {
  out.addFile(name, Buffer.from(content, 'utf8'));
}
out.writeZip(deckPath);
console.log('mutated deck.pptx: slide2 and slide3 edited, slide1 only noise');
fs.rmSync(tmp, { recursive: true, force: true });
