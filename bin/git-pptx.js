#!/usr/bin/env node
'use strict';

const commands = require('../lib/commands');

async function main() {
  const args = process.argv.slice(2);
  try {
    await commands.run(args);
  } catch (e) {
    console.error('git-pptx: ' + e.message);
    process.exitCode = 1;
  }
}

main();
