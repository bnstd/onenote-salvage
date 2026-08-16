#!/usr/bin/env python3
"""Render OneNote's handwriting to SVG, into an extract that already exists.

    onenote2md-ink --source ~/notebooks --out ~/markdown

A targeted second pass. Sections are chosen by reading the extract back — a page carrying
`handwriting: true` in its frontmatter is the marker `extract` wrote — so only sections that
contain ink are re-parsed, and the pass can be run any time after the first.

## The stroke encoding

`extract` records *where* the strokes are (jcid `0x00020047`, path in prid `0x340b`). This is
how they decode.

**Multi-byte signed integers.** Accumulate the low 7 bits of each byte until one arrives with
the top bit clear. The LSB of the result is the sign, the rest is the magnitude.

**The first value is a count** of the values that follow.

**Three dimensions: x, y, pressure.** The published description is "all of the X values, then
all of the Y values, then any other dimensions the stroke happens to record", which makes a
two-way split look reasonable. It is not, and the failure is not obvious: a wrong split
produces plausible handwriting with the pen apparently dragged across the page between every
letter, rather than noise. `_dimensions()` therefore verifies the split per stroke instead of
trusting the constant, scoring each candidate on whether the resulting path has steps far
larger than its median. The format allows other dimensions, and a wrong split is invisible in
the numbers alone.

**Coordinates are differentials**, accumulated from an absolute first value.
"""

import argparse
import os
import re
import statistics
import sys
from collections import defaultdict
from contextlib import contextmanager
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from onenote2md.extract import (  # noqa: E402
    GUID, Encrypted, _patched_pyonenote, _safe, _sections, parse_pages)


# A network share often needs the group to be able to edit what is written. That is a
# property of the destination, not of OneNote, so it is an option (`--umask 007`) and off by
# default. Never chmod a directory on such a share — it silently strips setgid.
SHARE_UMASK = None


@contextmanager
def share_umask():
    if SHARE_UMASK is None:
        yield
        return
    old = os.umask(SHARE_UMASK)
    try:
        yield
    finally:
        os.umask(old)

STROKE_JCID = 0x00020047
INK_PATH_PRID = 0x340b


def varints(data: bytes) -> list:
    out, acc, shift = [], 0, 0
    for b in data:
        acc |= (b & 0x7F) << shift
        if b & 0x80:
            shift += 7
            continue
        out.append(-(acc >> 1) if acc & 1 else (acc >> 1))
        acc, shift = 0, 0
    return out


def _accumulate(deltas):
    out, cur = [], 0
    for d in deltas:
        cur += d
        out.append(cur)
    return out


def _roughness(xs, ys) -> float:
    """Fraction of steps far larger than typical. A correct split scores ~0."""
    steps = [abs(xs[i] - xs[i - 1]) + abs(ys[i] - ys[i - 1]) for i in range(1, len(xs))]
    if not steps:
        return 1.0
    median = statistics.median(steps) or 1
    return sum(1 for s in steps if s > median * 20) / len(steps)


def _dimensions(rest) -> int:
    """How many dimensions this stroke records. Three, unless the numbers disagree."""
    best, chosen = None, 3
    for d in (3, 2, 4, 5):
        if len(rest) % d:
            continue
        n = len(rest) // d
        score = _roughness(_accumulate(rest[:n]), _accumulate(rest[n:2 * n]))
        if best is None or score < best:
            best, chosen = score, d
        if score == 0.0 and d == 3:
            break                       # the overwhelming common case; stop early
    return chosen


def stroke_points(blob: bytes):
    v = varints(blob)
    if len(v) < 7:
        return []
    rest = v[1:]
    n = len(rest) // _dimensions(rest)
    return list(zip(_accumulate(rest[:n]), _accumulate(rest[n:2 * n])))


def collect_strokes(path: Path, OneDocment) -> dict:
    """-> {object-space guid: [stroke, ...]}, each stroke a list of (x, y)."""
    # Imported here, not at module scope: _patched_pyonenote replaces the module in
    # sys.modules, so a name bound earlier points at the unpatched class and isinstance
    # silently matches nothing.
    import pyOneNote.FileNode as FN

    by_space = defaultdict(list)
    with open(path, 'rb') as f:
        doc = OneDocment(f)
        nodes = []
        OneDocment.traverse_nodes(doc.root_file_node_list, nodes,
                                  ['ObjectDeclaration2RefCountFND'])
        for node in nodes:
            if not hasattr(node, 'propertySet'):
                continue
            if node.data.body.jcid.jcid != STROKE_JCID:
                continue
            m = GUID.search(str(node.data.body.oid))
            if not m:
                continue
            ps = node.propertySet.body
            for prid, data in zip(ps.rgPrids, ps.rgData):
                if prid.id != INK_PATH_PRID:
                    continue
                if not isinstance(data, FN.PrtFourBytesOfLengthFollowedByData):
                    continue
                pts = stroke_points(data.Data)
                if len(pts) > 1:
                    by_space[m.group(1)].append(pts)
    return by_space


def svg(strokes, width=1200) -> str:
    xs = [p[0] for s in strokes for p in s]
    ys = [p[1] for s in strokes for p in s]
    x0, y0 = min(xs), min(ys)
    w, h = max(1, max(xs) - x0), max(1, max(ys) - y0)
    pad = max(w, h) * 0.02
    stroke_w = max(w, h) / 400
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" '
        f'height="{int(width * (h + 2 * pad) / (w + 2 * pad))}" '
        f'viewBox="{x0 - pad} {y0 - pad} {w + 2 * pad} {h + 2 * pad}">',
        f'<rect x="{x0 - pad}" y="{y0 - pad}" width="{w + 2 * pad}" '
        f'height="{h + 2 * pad}" fill="#ffffff"/>']
    for s in strokes:
        d = 'M ' + ' L '.join(f'{x} {y}' for x, y in s)
        parts.append(f'<path d="{d}" fill="none" stroke="#141414" '
                     f'stroke-width="{stroke_w:.2f}" stroke-linecap="round" '
                     f'stroke-linejoin="round"/>')
    parts.append('</svg>')
    return '\n'.join(parts)


_BANNER = re.compile(
    r'> \*\*Content here cannot be extracted\*\*.*?`_INK PAGES TO EXPORT\.md`\.\n\n',
    re.S)


def _link_into_page(md: Path, rel: str, strokes: int):
    """Replace the 'cannot be extracted' banner with the rendering, and say so up top."""
    text = md.read_text(errors='replace')
    if f'![[{rel}]]' in text:
        return False
    text = _BANNER.sub('', text)
    head, sep, body = text.partition('\n---\n')
    if not sep:
        return False
    if 'ink_strokes:' not in head:
        # `partition` strips the newline that ended the last frontmatter line, so put it back
        # before appending — otherwise the new key runs onto the end of `note:`.
        head = head.rstrip('\n') + f'\nink_strokes: {strokes}\nink_rendered: true'
    # Under the title, not above it: the H1 opens the page.
    body = body.lstrip('\n')
    m = re.match(r'(# .*\n)', body)
    figure = f'\n> Handwriting, rendered from the stroke data.\n\n![[{rel}]]\n'
    body = (m.group(1) + figure + body[m.end():]) if m else (figure + body)
    md.write_text(head + sep + '\n' + body)
    return True


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('--source', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--work', default='/tmp/onenote2md-ink')
    ap.add_argument('--exclude', action='append', default=[], metavar='GLOB')
    ap.add_argument('--umask', type=lambda v: int(v, 8), metavar='OCTAL',
                    help='umask for files written, e.g. 007 when a group needs write access '
                         'on a network share. Default: leave the process umask alone')
    args = ap.parse_args()
    if args.umask is not None:
        globals()['SHARE_UMASK'] = args.umask

    src, out, work = Path(args.source), Path(args.out), Path(args.work)
    work.mkdir(parents=True, exist_ok=True)
    OneDocment = _patched_pyonenote()

    # Which sections to bother with: the extract already records where handwriting was found.
    wanted = set()
    for md in out.rglob('*.md'):
        if md.name.startswith('_'):
            continue
        head = md.read_text(errors='replace').split('\n---\n', 1)[0]
        if 'handwriting: true' in head:
            wanted.add((md.parent.parent.name, md.parent.name))
    print(f'{len(wanted)} sections contain handwriting\n')

    from fnmatch import fnmatch
    pages_done = strokes_done = 0
    with share_umask():
        for pkg in sorted(src.rglob('*.onepkg')):
            if any(fnmatch(pkg.name, g) for g in args.exclude):
                continue
            notebook = _safe(pkg.stem, 'Notebook')
            if not any(nb == notebook for nb, _ in wanted):
                continue
            for section, tmp in _sections(pkg, work):
                sec = _safe(section, 'Section')
                if (notebook, sec) not in wanted:
                    continue
                sec_dir = out / notebook / sec
                try:
                    pages, _files, _note = parse_pages(Path(str(tmp)), OneDocment)
                except (Encrypted, Exception) as e:
                    print(f'  {sec:<40} skipped — {type(e).__name__}')
                    continue
                strokes_by_space = collect_strokes(Path(str(tmp)), OneDocment)
                if not strokes_by_space:
                    continue
                n_pages = n_strokes = 0
                for i, page in enumerate(pages, 1):
                    strokes = strokes_by_space.get(page['space'])
                    if not strokes:
                        continue
                    title = _safe(page['title'], f'Untitled page {i}')
                    md = sec_dir / f'{i:03d} {title}.md'
                    if not md.exists():
                        continue
                    fdir = sec_dir / f'{i:03d} {title}_files'
                    fdir.mkdir(exist_ok=True)
                    (fdir / 'handwriting.svg').write_text(svg(strokes))
                    if _link_into_page(md, f'{fdir.name}/handwriting.svg', len(strokes)):
                        n_pages += 1
                        n_strokes += len(strokes)
                pages_done += n_pages
                strokes_done += n_strokes
                print(f'  {notebook}/{sec:<34} {n_pages:>3} pages  {n_strokes:>6} strokes')

    print(f'\nrendered {strokes_done:,} strokes onto {pages_done} pages')


if __name__ == '__main__':
    main()
