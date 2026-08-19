'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFile, execFileSync } = require('child_process');

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
  // Cheap check via the App Paths registry key / known file paths. Never
  // launch PowerPoint just to detect it: a cold start is slow, and connecting
  // to an already-running instance only to quit it would close the user's deck.
  const candidates = [
    'C:\\Program Files\\Microsoft Office\\Root\\Office16\\POWERPNT.EXE',
    'C:\\Program Files (x86)\\Microsoft Office\\Root\\Office16\\POWERPNT.EXE',
  ];
  try {
    const out = execFileSync('powershell.exe', ['-NoProfile', '-Command',
      "if(Get-ItemProperty 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\POWERPNT.EXE' -ErrorAction SilentlyContinue){'Y'}"
    ], { encoding: 'utf8' });
    cachedPowerPoint = /Y/.test(out) || candidates.some((p) => fs.existsSync(p));
  } catch (e) {
    cachedPowerPoint = candidates.some((p) => fs.existsSync(p));
  }
  return cachedPowerPoint;
}

// In-place progress that rewrites a single line ("rendering slide k/N") instead
// of emitting one line per slide, so a large render does not flood the
// terminal. Used by both renderers so progress looks identical everywhere.
function makeProgress(total) {
  let k = 0;
  let shown = false;
  return {
    tick() {
      k++;
      const s = '  rendering slide ' + k + '/' + total;
      process.stderr.write(process.stderr.isTTY
        ? '\r\x1b[2m' + s + '\x1b[0m\x1b[K'
        : '\r' + s);
      shown = true;
    },
    end() {
      if (shown) process.stderr.write('\n');
    },
  };
}

// pairs: [{label, position}]. `position` is the real 1-based slide-show
// position to export (see deck.js#slidePartToPosition); `label` is only used
// to name the output file, so it can stay the slide part's own number even
// though that number is unrelated to `position`.
function renderLibreOffice(deckPath, outDir, pairs, ext, onProgress) {
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
    // Give headless soffice its own isolated user profile. This mirrors the
    // PowerPoint rule of never touching a running editor: a shared default
    // profile would lock against (and be blocked by) an open LibreOffice GUI,
    // so rendering never interrupts it.
    const profileDir = path.join(tmp, 'profile');
    const userInstallation = 'file:///' + profileDir.replace(/\\/g, '/');
    if (onProgress) onProgress('converting to PDF...');
    execFileSync(soffice, [
      '--headless', '--convert-to', 'pdf', '--outdir', pdfDir,
      '-env:UserInstallation=' + userInstallation, path.resolve(deckPath),
    ], { stdio: 'ignore' });
    const pdfPath = path.join(pdfDir, path.basename(deckPath, path.extname(deckPath)) + '.pdf');
    if (!fs.existsSync(pdfPath)) {
      throw new Error('LibreOffice failed to convert ' + deckPath + ' to PDF');
    }
    const fmtFlag = ext === 'png' ? '-png' : '-jpeg';
    const progress = makeProgress(pairs.length);
    pairs.forEach(({ label, position }) => {
      execFileSync(pdftoppm, [
        fmtFlag, '-r', '110', '-f', String(position), '-l', String(position), '-singlefile',
        pdfPath, path.join(outDir, String(label)),
      ], { stdio: 'ignore' });
      progress.tick();
    });
    progress.end();
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

// The single place that owns the PowerPoint COM lifecycle. It opens the deck
// hidden (WithWindow=false), exports the requested slides, closes the
// presentation without saving changes, and only quits PowerPoint when it
// started the instance itself. A shared (already-running) PowerPoint is never
// quit, so an open editor survives a render. Every render path goes through
// this one script, so there is no way to touch PowerPoint that can kill it.
function renderPowerPointScript(ext) {
  return [
    'param([string]$Deck, [string]$OutDir, [string]$Pairs)',
    "$ErrorActionPreference = 'Stop'",
    '$Running = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)',
    '$ppt = New-Object -ComObject PowerPoint.Application',
    '$ppt.DisplayAlerts = 1',
    'try {',
    '  $pres = $ppt.Presentations.Open($Deck, $false, $false, $false)',
    '  # Preserve the slide aspect ratio instead of forcing a fixed 1280x720.',
    '  $w = $pres.PageSetup.SlideWidth',
    '  $h = $pres.PageSetup.SlideHeight',
    '  $scale = 1280.0 / [Math]::Max($w, $h)',
    '  $ew = [int][Math]::Round($w * $scale)',
    '  $eh = [int][Math]::Round($h * $scale)',
    '  foreach ($pair in ($Pairs -split ",")) {',
    '    $fields = $pair -split ":"',
    '    $label = $fields[0]',
    '    $position = [int]$fields[1]',
    '    Write-Output ("rendering slide " + $label)',
    "    $name = $label + '.' + '" + ext + "'",
    "    $pres.Slides.Item($position).Export((Join-Path $OutDir $name), '" + ext.toUpperCase() + "', $ew, $eh)",
    '  }',
    '  $pres.Saved = $true',
    '  $pres.Close()',
    '} finally {',
    '  if (-not $Running) { $ppt.Quit() }',
    '}',
  ].join('\n');
}

async function renderPowerPoint(deckPath, outDir, pairs, ext, onProgress) {
  fs.mkdirSync(outDir, { recursive: true });
  if (onProgress) onProgress('opening presentation...');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'gpptx-render-'));
  const ps = path.join(tmp, 'render.ps1');
  fs.writeFileSync(ps, renderPowerPointScript(ext));
  // A long single COM session can drop with RPC 0x800706BA on large decks, so
  // each batch runs in its own session and is retried on transient failures.
  const CHUNK = 20;
  const progress = makeProgress(pairs.length);
  try {
    for (let i = 0; i < pairs.length; i += CHUNK) {
      const chunk = pairs.slice(i, i + CHUNK).map((p) => p.label + ':' + p.position).join(',');
      for (let attempt = 0; ; attempt++) {
        try {
          await runPowerShell(ps, deckPath, outDir, chunk, () => progress.tick());
          break;
        } catch (e) {
          if (attempt >= 2) throw e;
        }
      }
    }
  } finally {
    progress.end();
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

// Runs the render.ps1 PowerShell script and feeds each "rendering slide N" line
// it prints on stdout to `onSlide`. Capturing stdout (instead of inheriting it)
// lets this tool draw a single in-place progress line instead of one line per
// slide.
function runPowerShell(ps, deckPath, outDir, pairArg, onSlide) {
  return new Promise((resolve, reject) => {
    const child = execFile('powershell.exe', [
      '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', winPath(ps),
      '-Deck', winPath(deckPath), '-OutDir', winPath(outDir), '-Pairs', pairArg,
    ], (err) => (err ? reject(err) : resolve()));
    if (child.stdout) {
      let buf = '';
      child.stdout.on('data', (d) => {
        buf += String(d);
        let nl;
        while ((nl = buf.indexOf('\n')) !== -1) {
          const line = buf.slice(0, nl);
          buf = buf.slice(nl + 1);
          if (/^rendering slide /.test(line)) onSlide();
        }
      });
    }
  });
}

// Renders the given slide part/position pairs of a pptx to preview images
// (<label>.jpg, ...) in outDir. Only the listed pairs are rendered. Callers
// resolve `label` (the slide part's own number) to `position` (its real
// 1-based place in the slide show) via deck.js#slidePartToPosition — the two
// are only the same by coincidence for decks whose parts were never split out
// of a larger deck.
async function renderSlides(deckPath, outDir, pairs, format, renderer, onProgress) {
  const ext = format === 'png' ? 'png' : 'jpg';
  let backend = renderer || 'auto';
  if (backend === 'auto') {
    backend = (process.platform === 'win32' && hasPowerPoint()) ? 'powerpoint' : 'libreoffice';
  }
  if (backend === 'powerpoint') await renderPowerPoint(deckPath, outDir, pairs, ext, onProgress);
  else if (backend === 'libreoffice') renderLibreOffice(deckPath, outDir, pairs, ext, onProgress);
  else throw new Error('unknown renderer: ' + backend + ' (auto|powerpoint|libreoffice)');
}

module.exports = { renderSlides };
