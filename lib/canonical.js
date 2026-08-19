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
  return serializeNode(doc.documentElement);
}

module.exports = { canonicalize };
