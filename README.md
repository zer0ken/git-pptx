# git-pptx

[한국어 (Korean)](README.ko.md)

Version-control PowerPoint (`.pptx`) files per slide as a git-friendly directory. It is a standalone tool with no coupling to git or GitHub operations: `decomp`/`comp` convert between a pptx and the git-pptx directory format. Publish the result with ordinary git.

## git-pptx format

`decomp` of `a.pptx` produces an `a.git-pptx/` directory that sits next to the file:

```
a.git-pptx/
  previews/   1.jpg, 2.jpg, ...   per-slide previews (derived)
  pptx/       the pptx in unpacked form
```

- `decomp` is a lossless unzip; relationship IDs and `[Content_Types].xml` are not rebuilt, so no data is lost.
- `comp` re-zips `pptx/` (previews and VCS files excluded).
- Change detection is based on XML canonicalization: it ignores reserialization noise (attribute order, entity encoding, empty-element notation, namespace prefixes) and detects only slides whose real content changed. Element order (z-order) is preserved.

## Install

```bash
git clone https://github.com/zer0ken/git-pptx.git
cd git-pptx
npm link
```

`npm link` exposes the `git-pptx` executable on PATH.

## Usage

```bash
# a.pptx -> a.git-pptx/ (update changed slides, render previews)
git-pptx decomp a.pptx

# a.git-pptx/ -> a.pptx
git-pptx comp a.git-pptx a.pptx

# show changed slides without writing anything
git-pptx diff a.pptx a.git-pptx
```

Options:

- `--no-preview`: skip preview rendering
- `--format png`: render previews as PNG instead of JPG (default)
- `--renderer auto|powerpoint|libreoffice`: choose the preview renderer (default `auto`)

## Behavior

| Command | What it does |
|---|---|
| `git-pptx decomp a.pptx` | creates (or incrementally updates) `a.git-pptx/`. If a split already exists, only changed files are updated and previews are re-rendered only for changed slides. |
| `git-pptx comp <dir> <out.pptx>` | zips `pptx/` back into a `.pptx` |
| `git-pptx diff <deck.pptx> <dir>` | lists changed slides |

## Discussing on GitHub

Committing `previews/1.jpg, 2.jpg, ...` lets GitHub render image previews and PR image diffs, so slides can be referenced by number and discussed with inline comments. Add `*.pptx` to `.gitignore` and push only the `a.git-pptx/` directory.

## Limitations

- `docProps/core.xml` (timestamp metadata) changes on every save and is excluded from change detection.
- Canonicalization ignores reserialization formatting noise but does not yet handle relationship ID renumbering (`r:id`) or default-value materialization; an unchanged slide may appear changed if it is renumbered.
- Preview rendering is cross-platform: `libreoffice` (LibreOffice headless + poppler, all OSes) or `powerpoint` (PowerPoint COM, Windows). The default `auto` picks PowerPoint on Windows when available, else LibreOffice. Previews differ between renderers, so a team should standardize on one via `--renderer`.

## Development

- `scripts/make-fixture.js`: generate a minimal test pptx
- `scripts/mutate.js`: mutate a deck with reserialization noise and real edits
