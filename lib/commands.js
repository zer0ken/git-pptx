'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');

const configMod = require('./config');
const deckMod = require('./deck');
const { diffDirs } = require('./diff');

function git(root, ...args) {
  return execFileSync('git', args, {
    cwd: root,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  });
}

function tmpdir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'pit-'));
}

function slideLabel(rel) {
  const m = rel.match(/^ppt\/slides\/slide(\d+)\.xml$/);
  if (m) return 'slide ' + m[1];
  return rel;
}

function readArgs(args) {
  const flags = {};
  const positional = [];
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a.startsWith('--')) {
      const key = a.slice(2);
      const val = args[i + 1];
      flags[key] = val === undefined || val.startsWith('--') ? true : val;
      if (flags[key] !== true) i++;
    } else {
      positional.push(a);
    }
  }
  return { flags, positional };
}

function ensureInit(root) {
  const cfg = configMod.load(root);
  if (!cfg) {
    throw new Error('not initialized. run: git pit init');
  }
  return cfg;
}

async function cmdInit(root, args) {
  if (configMod.load(root)) {
    throw new Error('already initialized');
  }
  const { flags } = readArgs(args);
  const deck = flags.deck || 'deck.pptx';
  const slidesDir = flags['slides-dir'] || 'slides';
  const upstream = flags.upstream || null;

  const cfg = { deck, slidesDir };
  if (upstream) cfg.upstream = upstream;
  configMod.save(root, cfg);

  // Do not track the binary deck; only the split directory is tracked.
  const giPath = path.join(root, '.gitignore');
  const gi = fs.existsSync(giPath) ? fs.readFileSync(giPath, 'utf8') : '';
  if (!gi.split('\n').includes(deck)) {
    fs.appendFileSync(giPath, deck + '\n');
  }

  // Keep the split directory byte-exact so git never normalizes line endings,
  // which would otherwise show every slide as changed on each push.
  const gaPath = path.join(root, '.gitattributes');
  const rule = slidesDir + '/** -text';
  const ga = fs.existsSync(gaPath) ? fs.readFileSync(gaPath, 'utf8') : '';
  if (!ga.includes(rule)) {
    fs.appendFileSync(gaPath, rule + '\n');
  }

  const deckPath = path.join(root, deck);
  const slidesPath = path.join(root, slidesDir);

  if (fs.existsSync(deckPath)) {
    fs.mkdirSync(slidesPath, { recursive: true });
    deckMod.unzipToDir(deckPath, slidesPath);
    git(root, 'add', slidesPath, '.pit', '.gitignore', '.gitattributes');
    git(root, 'commit', '-m', 'pit: initial split of ' + deck, '--allow-empty');
    console.log('split ' + deck + ' into ' + slidesDir + '/');
  }

  if (upstream) {
    try {
      git(root, 'remote', 'add', 'origin', upstream);
      git(root, 'push', '-u', 'origin', 'HEAD');
      console.log('pushed to ' + upstream);
    } catch (e) {
      console.log('remote not pushed: ' + e.message.split('\n')[0]);
    }
  }
  console.log('initialized.');
}

async function cmdPush(root, args) {
  const cfg = ensureInit(root);
  const deckPath = path.join(root, cfg.deck);
  const slidesPath = path.join(root, cfg.slidesDir);
  if (!fs.existsSync(deckPath)) {
    throw new Error('deck not found: ' + cfg.deck);
  }
  fs.mkdirSync(slidesPath, { recursive: true });

  const tmp = tmpdir();
  try {
    deckMod.unzipToDir(deckPath, tmp);
    const changed = diffDirs(tmp, slidesPath);
    if (changed.length === 0) {
      console.log('no changes');
      return;
    }
    for (const rel of changed) {
      const src = path.join(tmp, rel);
      const dst = path.join(slidesPath, rel);
      if (fs.existsSync(src)) {
        fs.mkdirSync(path.dirname(dst), { recursive: true });
        fs.copyFileSync(src, dst);
      } else {
        fs.rmSync(dst, { force: true });
      }
    }
    git(root, 'add', slidesPath, '.pit');
    const labels = changed.map(slideLabel).join(', ');
    git(root, 'commit', '-m', 'pit: update ' + labels, '--allow-empty');
    console.log('updated ' + changed.length + ' file(s): ' + labels);

    if (cfg.upstream) {
      try {
        git(root, 'push');
        console.log('pushed.');
      } catch (e) {
        console.log('committed locally; push manually: git push');
      }
    }
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

function hasUpstream(root) {
  try {
    git(root, 'rev-parse', '--abbrev-ref', '--symbolic-full-name', '@{u}');
    return true;
  } catch (e) {
    return false;
  }
}

async function cmdPull(root, args) {
  const cfg = ensureInit(root);
  const slidesPath = path.join(root, cfg.slidesDir);
  if (hasUpstream(root)) {
    try {
      git(root, 'pull');
    } catch (e) {
      throw new Error('git pull failed: ' + e.message.split('\n')[0]);
    }
  } else {
    console.log('no upstream configured; rebuilding from local ' + cfg.slidesDir + '/');
  }
  const deckPath = path.join(root, cfg.deck);
  deckMod.zipDirToFile(slidesPath, deckPath);
  console.log('rebuilt ' + cfg.deck + ' from ' + cfg.slidesDir + '/');
}

async function cmdStatus(root, args) {
  const cfg = ensureInit(root);
  const deckPath = path.join(root, cfg.deck);
  const slidesPath = path.join(root, cfg.slidesDir);
  if (!fs.existsSync(deckPath)) {
    throw new Error('deck not found: ' + cfg.deck);
  }
  fs.mkdirSync(slidesPath, { recursive: true });
  const tmp = tmpdir();
  try {
    deckMod.unzipToDir(deckPath, tmp);
    const changed = diffDirs(tmp, slidesPath);
    if (changed.length === 0) {
      console.log('in sync');
      return;
    }
    for (const rel of changed) console.log('* ' + slideLabel(rel));
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

const HELP = `git pit - version-control PowerPoint decks as per-slide files

Usage:
  git pit init [--deck FILE] [--slides-dir DIR] [--upstream URL]
      split deck.pptx into DIR/ and record config
  git pit push        update changed slides from the local deck and commit
  git pit pull        rebuild the deck from the remote split directory
  git pit status      show which slides differ from the committed state
`;

async function run(root, cmd, args) {
  switch (cmd) {
    case 'init': return cmdInit(root, args);
    case 'push': return cmdPush(root, args);
    case 'pull': return cmdPull(root, args);
    case 'status': return cmdStatus(root, args);
    case 'help':
    case '--help':
    case '-h': console.log(HELP); return;
    default: throw new Error('unknown command: ' + cmd + '\n' + HELP);
  }
}

module.exports = { run };
