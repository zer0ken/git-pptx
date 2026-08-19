# git-pptx

[한국어 (Korean)](README.ko.md)

Version-control PowerPoint (`.pptx`) files per slide as a git-friendly directory. It is standalone and does not couple to git or GitHub operations: `decomp`/`comp` convert between a pptx and the `a.git-pptx/` directory format. Publish the result with ordinary git.

## Install

```bash
git clone https://github.com/zer0ken/git-pptx.git
cd git-pptx
npm link
```

`npm link` exposes the `git-pptx` executable on PATH.

## Usage

```bash
git-pptx decomp a.pptx              # a.pptx -> a.git-pptx/ (update changed slides, render previews)
git-pptx comp a.git-pptx out.pptx   # a.git-pptx/ -> pptx
git-pptx diff a.pptx a.git-pptx     # show changed slides without writing
```

`decomp` unzips `a.pptx` into `a.git-pptx/` (layout below), renames numbered parts to the dense 1..N sequence PowerPoint uses, and detects only slides whose real content changed. `comp` re-zips `pptx/` back.

```
a.git-pptx/
  previews/   1.jpg, 2.jpg, ...   per-slide previews (derived)
  pptx/       the pptx in unpacked form
```

Options:

- `--no-preview`: skip preview rendering
- `--format png`: previews as PNG instead of JPG (default)
- `--renderer auto|powerpoint|libreoffice`: preview renderer (default `auto`)
- `--no-normalize`: keep the part names the deck came with

## Output and rendering

- Results go to stdout (bold), progress to stderr (dim); plain text when not a terminal. `decomp`/`diff` summarize changed slides.
- Rendering never touches an open editor: previews are rendered from a temp copy, a running PowerPoint is never quit, and LibreOffice runs in an isolated profile.
- Renderer: `powerpoint` (COM, Windows) or `libreoffice` (headless + poppler, all OSes). `auto` picks PowerPoint on Windows when available. Previews differ between renderers; standardize with `--renderer`.

## Discussing on GitHub

Commit `previews/` so GitHub renders image previews and PR image diffs. Add `*.pptx` to `.gitignore` and push only `a.git-pptx/`.

## Limitations

- `docProps/core.xml`, `docProps/app.xml`, `docProps/thumbnail.jpeg` are regenerated on every save and excluded from change detection (still written so `comp` stays valid).
- Canonicalization ignores reserialization noise but not relationship ID renumbering (`r:id`); an unchanged part can look changed after the first PowerPoint save of a foreign-written deck.
- SVG+raster decks may have a few media parts renamed once on the first save.
