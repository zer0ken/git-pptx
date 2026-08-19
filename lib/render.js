'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');

function winPath(p) {
  return path.resolve(p).replace(/\//g, '\\');
}

// Renders the given 1-based slide indices of a pptx to preview images
// (1.jpg, 2.jpg, ...) in outDir using installed PowerPoint via COM.
// Only the listed indices are rendered, so unchanged slides are not re-rendered.
function renderSlides(deckPath, outDir, indices, format) {
  const ext = format === 'png' ? 'png' : 'jpg';
  fs.mkdirSync(outDir, { recursive: true });
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'gpptx-render-'));
  const ps = path.join(tmp, 'render.ps1');
  fs.writeFileSync(ps, [
    'param([string]$Deck, [string]$OutDir, [string]$Indices)',
    "$ErrorActionPreference = 'Stop'",
    '$ppt = New-Object -ComObject PowerPoint.Application',
    'try {',
    '  $pres = $ppt.Presentations.Open($Deck, $true, $false, $false)',
    '  foreach ($i in ($Indices -split ",")) {',
    "    $name = ([string]$i) + '.' + '" + ext + "'",
    "    $pres.Slides.Item([int]$i).Export((Join-Path $OutDir $name), '" + ext.toUpperCase() + "', 1280, 720)",
    '  }',
    '  $pres.Close()',
    '} finally {',
    '  $ppt.Quit()',
    '}',
  ].join('\n'));
  try {
    execFileSync('powershell.exe', [
      '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', winPath(ps),
      '-Deck', winPath(deckPath), '-OutDir', winPath(outDir), '-Indices', indices.join(','),
    ], { stdio: 'inherit' });
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

module.exports = { renderSlides };
