# AGENTS.md

Guidance for an AI coding agent driving this tool.

## Run order

```bash
onenote2md     --source <folder of .onepkg/.one> --out <markdown tree>
onenote2md-ink --source <same>                   --out <same>
```

1. **Extract.** One markdown file per page, images beside it, XPS printouts converted to
   PDF, a per-section report of what was found and what was skipped.
2. **Ink.** Renders handwriting to SVG and links it into the pages. Reads the extract to
   find which sections contain ink, so it only re-parses those. Run any time after step 1.

`--resume` skips sections that already produced output. A failed section produces no
output and is retried automatically, so nothing needs deleting before a re-run.

Add `--umask 007` when writing to a share where a group needs write access.

## Output shape

```
<out>/<notebook>/<section>/NNN <page title>.md
<out>/<notebook>/<section>/NNN <page title>_files/<attachments>
```

Frontmatter carries `title`, `notebook`, `section`, `created`, `images`, `typed_chars`,
`ocr_chars`, and `handwriting: true` where ink was found. Image links are
`![[wiki-style]]`.

## Things that will catch you out

**An empty section may be an encrypted one.** A password-protected section parses without
error and yields nothing. The tool raises `Encrypted` and reports `ENCRYPTED` rather than
returning an empty result, so check the report before concluding a section had no content.

**Unlocking is not removing a password.** If a section is still reported as encrypted after
the user says they unlocked it, that is why — unlocking decrypts for that session only, and
an export taken afterwards still contains the encrypted bytes. They need Password
Protection → **Remove Password**, then a fresh export.

**`typed_chars` is not "text the user wrote".** Text on an image node is OneNote's OCR of
that image, so a scanned document can report a six-figure character count. Use `ocr_chars`
to tell documents from notes. Beware the next layer too: pasted content — clipped articles,
forwarded email — counts as typed by any mechanical test, so `typed_chars` alone will not
identify original writing.

**Attachments are duplicated before deduplication.** An image node appears once per revision
that touched it. If you are counting or reporting what an archive contains, deduplicate by
content hash first or you will overstate it substantially.

**OCR text is per attachment, not per page.** A converted printout can run to dozens of
pages. Writing OCR block N onto page N misplaces everything after the first multi-page
attachment.

**Verify ink by looking at it.** A wrong stroke decode produces plausible handwriting with
the pen apparently dragged between letters, not obvious garbage. Render a page and open it.

**Do not re-encode images.** The archive may hold the only copy of a scan at that
resolution.

## If you are building the layer above this

The tool extracts; it does not decide where anything belongs.

**Classify by section, not by page.** A section name usually says what it is; an individual
page title often does not. A table of section-to-destination mappings that the user reviews
beats several thousand per-page inferences, and pasted correspondence otherwise classifies
as personal writing.

**A page's scans are one document.** Filing them individually produces orphaned images that
cannot be related back to anything. Merge them, and carry OneNote's OCR into the result as
an invisible text layer to keep it searchable.

## Working on the tool itself

`pyOneNote` is patched **at import time** (`_PATCHES`, `_DOC_PATCHES`, `_patch_module`), not
forked. Each patch anchors on an exact excerpt of upstream source and is **fatal if it stops
matching**, so an upstream change cannot leave the parser silently unguarded. The dependency
is pinned to `pyOneNote==0.0.2` for the same reason. If a patch stops applying, read the
upstream diff before touching the anchor.

The patches exist because values read from the file are used unchecked upstream. When adding
one, bound the value **by what the format permits** rather than by a guess about plausible
sizes, and count what you skip — `SKIPPED` reports it so a dropped object is not silently
assumed to be zero.

`--max-memory-gb` is not a tuning knob. It converts a kernel OOM kill, which loses the whole
run and exits zero, into one `MemoryError` costing one section. Set it **below the machine's
real ceiling**, which is not its installed RAM; a cap above the kill point does nothing.
