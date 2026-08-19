#!/usr/bin/env node
'use strict';

const { execSync } = require('child_process');
const commands = require('../lib/commands');

function gitRoot() {
  try {
    return execSync('git rev-parse --show-toplevel', { encoding: 'utf8' }).trim();
  } catch (e) {
    throw new Error('not inside a git repository');
  }
}

async function main() {
  const args = process.argv.slice(2);
  const cmd = args[0] || 'help';
  const rest = args.slice(1);
  const root = gitRoot();
  try {
    await commands.run(root, cmd, rest);
  } catch (e) {
    console.error('pit: ' + e.message);
    process.exitCode = 1;
  }
}

main();
