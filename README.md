# onenote2md

Export OneNote notebooks to markdown, with images, printouts and handwriting.

## Install

```bash
pip install onenote2md
```

Requires Python 3.9+.

## Usage

```bash
onenote2md     --source ~/notebooks --out ~/markdown
onenote2md-ink --source ~/notebooks --out ~/markdown
```

`--source` is a folder containing `.onepkg` exports (File → Export → Notebook → OneNote
Package) or loose `.one` section files. Subfolders are searched.

`onenote2md-ink` is a second pass that renders handwriting to SVG and links it into
the pages. It reads the extract to find which sections contain ink, so it only re-parses
those. Run it any time after the first pass.

### Options

For `onenote2md`. The ink pass takes `--source`, `--out`, `--work`, `--exclude` and
`--umask` only.

| Option | |
|---|---|
| `--source PATH` | Folder searched for `*.onepkg` and `*.one` |
| `--out PATH` | Where to write the markdown |
| `--resume` | Skip sections that already produced output; retries failures |
| `--exclude GLOB` | Skip matching source files (repeatable) |
| `--convert-xps` | Convert `.xps` files already under `--out` to PDF; no extraction |
| `--workers N` | Parallel processes for `--convert-xps` (default 5) |
| `--work PATH` | Scratch directory for unpacking (default `/tmp/onenote2md`) |
| `--max-memory-gb N` | Address-space cap, so a corrupt section fails alone (default 2.0) |
| `--umask OCTAL` | Umask for written files, e.g. `007` for a group-writable share |

## Output

One markdown file per page, as `<out>/<notebook>/<section>/NNN <page title>.md`, with that
page's images in `NNN <page title>_files/` beside it and linked inline.

```markdown
---
title: Kitchen quotes
notebook: House
section: Renovation
created: 2019-04-12
source: onenote
images: 3
typed_chars: 412
ocr_chars: 8801
note: "3/3 images linked"
---
```

Image links use `![[wiki-style]]` syntax, which suits Obsidian and is simple to rewrite for
anything else.

`typed_chars` and `ocr_chars` are separated deliberately. Text stored on an image node is
OneNote's own OCR of that image, not something you typed, so a page of scans can otherwise
report a six-figure character count and look like a long note. Use `ocr_chars` to spot
documents masquerading as notes.

Each run also writes a report of what was found per section, including anything skipped.

## What it handles

- **Handwriting.** Ink is stored as vector strokes with no image to recover.
  `onenote2md-ink` decodes the strokes and renders them to SVG.
- **Images a plain parse loses.** A blob is declared once per revision that referenced it,
  each with a different identity, and only the last is normally kept — so every page
  pointing at an earlier one silently loses its image. In one section that was 174 of 543.
- **"Print to OneNote" output.** Printouts are stored as XPS, which little outside Windows
  opens. They are converted to PDF, and OneNote's OCR of them is preserved since neither
  form carries a text layer.
- **Encrypted sections.** A password-protected section parses without error and yields
  nothing, which looks exactly like an empty section. It is reported as `ENCRYPTED`.
- **Corrupt files.** Lengths, counts and offsets read from the file are bounded by what the
  format permits, and anything skipped is counted and reported rather than assumed to be
  zero.

## Limitations

- **Password-protected sections** cannot be read. In OneNote, use Password Protection →
  **Remove Password** — unlocking is not enough, it decrypts for the session only and an
  export taken afterwards still contains the encrypted bytes — then export again.
- **The undocumented `638de92f` format** is not supported. Re-export the notebook as
  `.onepkg` and OneNote writes the documented one.
- **No OCR.** Where OneNote recognised text, that is preserved. Otherwise run something
  like [OCRmyPDF](https://github.com/ocrmypdf/OCRmyPDF) over the output.

## See also

[onenote.rs](https://github.com/msiemens/onenote.rs) reads the format more completely and
renders ink. If you are comfortable with Rust, start there.

`AGENTS.md` in this repo is written for driving the tool with an AI coding agent.

## Credits

Built on [pyOneNote](https://github.com/DissectMalware/pyOneNote) by Amirreza Niakanlahiji,
which does the revision-store parsing. Parser fixes are applied to its source at import time
and listed in [`NOTICE`](NOTICE); the dependency is pinned to `pyOneNote==0.0.2` because each
one anchors on an exact excerpt.

The ink encoding was reverse-engineered and published by
[Michael Siemens](https://m-siemens.de/blog/2026/05/decoding-onenote-s-file-format-secrets/).

## License

Apache-2.0, as pyOneNote is. See [`NOTICE`](NOTICE) for what has been changed.
