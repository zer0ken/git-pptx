'use strict';

const fs = require('fs');
const path = require('path');

function configPath(root) {
  return path.join(root, '.pit', 'config.json');
}

function defaults() {
  return { deck: 'deck.pptx', slidesDir: 'slides' };
}

function load(root) {
  const p = configPath(root);
  if (!fs.existsSync(p)) return null;
  return Object.assign(defaults(), JSON.parse(fs.readFileSync(p, 'utf8')));
}

function save(root, cfg) {
  const dir = path.join(root, '.pit');
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(configPath(root), JSON.stringify(cfg, null, 2) + '\n');
}

module.exports = { load, save, defaults, configPath };
