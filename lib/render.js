'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');

// Multi-backend preview rendering.
//   powerpoint  : PowerPoint COM (Windows only). Closest to what PowerPoint shows.
//   libreoffice : LibreOffice(headless) -> PDF -> pdftoppm. Works on all OSes.
// Default ('auto') picks powerpoint on Windows when available, else libreoffice.
// Use --renderer to force a backend.

function winPath(p) {
  return path.resolve(p).replace(/\//g, '\\');
}

function findTool(names) {
  for (const c of names) {
    try {
      execFileSync(c, ['--version'], { stdio: 'ignore' });
      return c;
    } catch (e) { /* try next */ }
  }
  return null;
}

function findPdftoppm() {
  // Some pdftoppm builds (e.g. MikTeX) reject --version under elevation, so
  // probe with -v and fall back to scanning the winget Poppler install dir.
  for (const c of ['pdftoppm', 'pdftoppm.exe']) {
    try {
      execFileSync(c, ['-v'], { stdio: 'ignore' });
      return c;
    } catch (e) { /* try next */ }
  }
  const base = path.join(process.env.LOCALAPPDATA || '', 'Microsoft', 'WinGet', 'Packages');
  if (fs.existsSync(base)) {
    for (const dir of fs.readdirSync(base)) {
      if (!dir.startsWith('oschwartz10612.Poppler_')) continue;
      const root = path.join(base, dir);
      for (const sub of fs.readdirSync(root)) {
        const p = path.join(root, sub, 'Library', 'bin', 'pdftoppm.exe');
        if (fs.existsSync(p)) return p;
      }
    }
  }
  return null;
}

function findSoffice() {
  const t = findTool(['soffice', 'soffice.exe', 'libreoffice']);
  if (t) return t;
  const win = [
    'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  ];
  for (const p of win) if (fs.existsSync(p)) return p;
  const mac = '/Applications/LibreOffice.app/Contents/MacOS/soffice';
  if (fs.existsSync(mac)) return mac;
  return null;
}

let cachedPowerPoint = null;
function hasPowerPoint() {
  if (cachedPowerPoint !== null) return cachedPowerPoint;
  if (process.platform !== 'win32') {
    cachedPowerPoint = false;
    return false;
  }
  try {
    const ps = "try { $p = New-Object -ComObject PowerPoint.Application; $p.Quit(); 'OK' } catch { 'NO' }";
    const out = execFileSync('powershell.exe', ['-NoProfile', '-Command', ps], { encoding: 'utf8' });
    cachedPowerPoint = /OK/.test(out);
  } catch (e) {
    cachedPowerPoint = false;
  }
  return cachedPowerPoint;
}

// pairs: [{label, position}]. `position` is the real 1-based slide-show
// position to export (see deck.js#slidePartToPosition); `label` is only used
// to name the output file, so it can stay the slide part's own number even
// though that number is unrelated to `position`.
function renderLibreOffice(deckPath, outDir, pairs, ext) {
  const soffice = findSoffice();
  if (!soffice) {
    throw new Error('LibreOffice not found. Install it to render previews: https://www.libreoffice.org');
  }
  const pdftoppm = findPdftoppm();
  if (!pdftoppm) {
    throw new Error('pdftoppm (Poppler) not found. Install poppler-utils to render previews.');
  }
  fs.mkdirSync(outDir, { recursive: true });
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'gpptx-render-'));
  try {
    const pdfDir = path.join(tmp, 'out');
    fs.mkdirSync(pdfDir, { recursive: true });
    execFileSync(soffice, [
      '--headless', '--convert-to', 'pdf', '--outdir', pdfDir, path.resolve(deckPath),
    ], { stdio: 'ignore' });
    const pdfPath = path.join(pdfDir, path.basename(deckPath, path.extname(deckPath)) + '.pdf');
    if (!fs.existsSync(pdfPath)) {
      throw new Error('LibreOffice failed to convert ' + deckPath + ' to PDF');
    }
    const fmtFlag = ext === 'png' ? '-png' : '-jpeg';
    for (const { label, position } of pairs) {
      execFileSync(pdftoppm, [
        fmtFlag, '-r', '110', '-f', String(position), '-l', String(position), '-singlefile',
        pdfPath, path.join(outDir, String(label)),
      ], { stdio: 'ignore' });
    }
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

function renderPowerPoint(deckPath, outDir, pairs, ext) {
  fs.mkdirSync(outDir, { recursive: true });
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'gpptx-render-'));
  const ps = path.join(tmp, 'render.ps1');
  const pairArg = pairs.map((p) => p.label + ':' + p.position).join(',');
  fs.writeFileSync(ps, [
    'param([string]$Deck, [string]$OutDir, [string]$Pairs)',
    "$ErrorActionPreference = 'Stop'",
    '$ppt = New-Object -ComObject PowerPoint.Application',
    'try {',
    '  $pres = $ppt.Presentations.Open($Deck, $true, $false, $false)',
    '  foreach ($pair in ($Pairs -split ",")) {',
    '    $fields = $pair -split ":"',
    '    $label = $fields[0]',
    '    $position = [int]$fields[1]',
    "    $name = $label + '.' + '" + ext + "'",
    "    $pres.Slides.Item($position).Export((Join-Path $OutDir $name), '" + ext.toUpperCase() + "', 1280, 720)",
    '  }',
    '  $pres.Close()',
    '} finally {',
    '  $ppt.Quit()',
    '}',
  ].join('\n'));
  try {
    execFileSync('powershell.exe', [
      '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', winPath(ps),
      '-Deck', winPath(deckPath), '-OutDir', winPath(outDir), '-Pairs', pairArg,
    ], { stdio: 'inherit' });
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

// Renders the given slide part/position pairs of a pptx to preview images
// (<label>.jpg, ...) in outDir. Only the listed pairs are rendered. Callers
// resolve `label` (the slide part's own number) to `position` (its real
// 1-based place in the slide show) via deck.js#slidePartToPosition — the two
// are only the same by coincidence for decks whose parts were never split out
// of a larger deck.
function renderSlides(deckPath, outDir, pairs, format, renderer) {
  const ext = format === 'png' ? 'png' : 'jpg';
  let backend = renderer || 'auto';
  if (backend === 'auto') {
    backend = (process.platform === 'win32' && hasPowerPoint()) ? 'powerpoint' : 'libreoffice';
  }
  if (backend === 'powerpoint') renderPowerPoint(deckPath, outDir, pairs, ext);
  else if (backend === 'libreoffice') renderLibreOffice(deckPath, outDir, pairs, ext);
  else throw new Error('unknown renderer: ' + backend + ' (auto|powerpoint|libreoffice)');
}

module.exports = { renderSlides };
