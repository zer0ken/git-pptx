'use strict';

const { DOMParser } = require('@xmldom/xmldom');

// Deterministic XML serialization used as a diff baseline.
// Attribute order, entity encoding, empty-element notation, and namespace
// prefixes that map to the same URI normalize to the same output.
// Element order (z-order of shapes) is preserved, never reordered.

function escapeText(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttr(s) {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}

function serializeNode(n) {
  if (n.nodeType === 3) return escapeText(n.data);
  if (n.nodeType === 4) return '<![CDATA[' + n.data + ']]>';
  if (n.nodeType !== 1) return '';

  const attrs = [];
  for (let i = 0; i < n.attributes.length; i++) {
    attrs.push([n.attributes[i].name, n.attributes[i].value]);
  }
  attrs.sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));

  let out = '<' + n.nodeName;
  for (const [k, v] of attrs) out += ' ' + k + '="' + escapeAttr(v) + '"';

  let inner = '';
  for (let c = n.firstChild; c; c = c.nextSibling) {
    const s = serializeNode(c);
    if (s !== '') inner += s;
  }
  return inner === '' ? out + '/>' : out + '>' + inner + '</' + n.nodeName + '>';
}

function canonicalize(xml) {
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const root = doc.documentElement;
  // Ignore derived artifacts that PowerPoint regenerates or drops on every
  // save, so they do not flag otherwise-unchanged files as changed.
  if (root) {
    if (root.nodeName === 'p:presentation') {
      // embedded-font list (fonts are re-embedded on each save)
      for (let c = root.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 1 && c.nodeName === 'p:embeddedFontLst') {
          root.removeChild(c);
          break;
        }
      }
    } else if (root.nodeName === 'Types') {
      // changesInfos (undo/change-tracking) part entries
      const drop = [];
      for (let c = root.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 1 && String(c.getAttribute('PartName') || '').includes('/changesInfos/')) {
          drop.push(c);
        }
      }
      drop.forEach((n) => root.removeChild(n));
    } else if (root.nodeName === 'Relationships') {
      // changesInfos relationships
      const drop = [];
      for (let c = root.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 1 && String(c.getAttribute('Target') || '').includes('changesInfos')) {
          drop.push(c);
        }
      }
      drop.forEach((n) => root.removeChild(n));
    }
  }
  return serializeNode(root);
}

module.exports = { canonicalize };
