"""Extract OneNote notebooks to markdown — notebook > section > page > content.

    onenote2md --source ~/notebooks --out ~/markdown

Reads `*.onepkg` (a Microsoft CAB) and loose `*.one` sections. Writes one markdown file per
page, named `NNN <page title>.md`, with that page's images beside it in `NNN <title>_files/`
and linked inline as `![[...]]`.

## The two on-disk formats

A `.one` file declares a format GUID in its header, and only one is documented:

    109add3f-911b-49f5-a5d0-1791edc8aed8   MS-ONESTORE. Parses.
    638de92f-a6d4-4bc1-9a36-b3fc2511a5b7   Undocumented. Does not.

Every field after the GUID belongs to a different layout, so a `638de92f` file parses as
nonsense rather than failing cleanly. Re-exporting the notebook as `.onepkg` (or Save As →
"OneNote 2010-2016 Section") writes it in the documented format.

## How the structure is recovered

Each page is its own *object space*, and every node carries that space's GUID in its
ExtendedGUID identity, so grouping the property list by GUID partitions a section into
pages. `jcidPageMetaData` gives `CachedTitleString` and `TopologyCreationTimeStamp`.

Images are matched by identity, not by order. `get_files()` is keyed by file-data-store
GUID, which nothing else references; each entry also records the identity that
`jcidImageNode.PictureContainer` points at. Order would be wrong even where the counts line
up, because `get_files()` also returns superseded images from version history that belong to
no current page. Those are reported as "unused (old revisions)".

A blob is declared once per revision that referenced it, each declaration carrying a
different identity, and upstream keeps only the last. Any page pointing at an earlier one
then resolves to nothing even though the bytes are present. `_DOC_PATCHES` accumulates every
identity and `parse_pages` resolves against all of them.

## Password-protected sections

A locked section parses without error and yields nothing: `get_properties()` returns an empty
list while `get_files()` still reports its blobs, because the object declarations are
encrypted and the file-data store is not. That is indistinguishable from an empty section, so
`parse_pages` raises `Encrypted` instead — the presence of `ObjectDataEncryptionKeyV2FNDX`
nodes with no object declarations is the tell. Sections in this state also report
`SKIPPED Nxfiledata`, since lengths read out of encrypted bytes fail the bounds checks.

There is no code fix. pyOneNote implements no decryption, and *unlocking* a section decrypts
it for that session only, so an export taken afterwards still carries the encrypted bytes.
The password must be removed (Password Protection → Remove Password) and the notebook
re-exported.

## The memory cap

A corrupt length field makes the parser ask for more memory than the machine will give, and
an unbounded process is then killed by the kernel — losing every notebook after the bad
section, with an exit code the shell reports as success. `--max-memory-gb` caps the address
space so the allocation raises `MemoryError`, which `write_section` already treats as an
unparseable section, costing one section instead of the run.

The cap must sit below the machine's real ceiling, which is not its RAM — a cgroup or
overcommit limit will kill the process well below the installed total. A cap above the kill
point does nothing at all. Measure before raising it.

## Resuming

`--resume` skips any section that already has output. A section's directory is created only
*after* its parse succeeds, so an existing directory means that section worked and anything
that failed is retried.

The ink checklist is rebuilt by reading the written markdown back (`_ink_rows`) rather than
from what the current run extracted, so resuming cannot produce a checklist missing pages an
earlier run found. The same applies to every file that claims to describe the corpus, so the
output is split by what each file is *for*:

    _extraction log.txt    the corpus, rebuilt from the frontmatter on disk (`_corpus_summary`)
    _extraction runs.txt   this run's transcript, appended under a dated header

Neither can be partial: the first is derived from the output itself, the second is only ever
added to. The rule generalises — **a file that claims to describe the whole extract is built by
reading the extract, never by collecting as you go.** Written the other way round, a resumed
run silently replaces the log with the handful of sections it did not skip, and it still looks
like a log.

## Ink

Handwriting is stored as vector strokes, not pixels. pyOneNote has no name for any ink jcid,
so `_patched_pyonenote()` labels those objects `Unknown(0x...)`; they then appear empty
because `PropertySet.get_properties` drops every property whose id it cannot name:

    propertyName = str(self.rgPrids[i])
    if propertyName != 'Unknown':

An object carrying only ink properties therefore formats as `{}`. The data is parsed and then
discarded. The two jcids that matter:

    jcid 0x00020047  stroke node           prid 0x340b holds the paths
    jcid 0x00060014  positioned container  OffsetFromParentHoriz/Vert

`onenote2md-ink` decodes and renders them; see that module for the encoding. This one
only records where the strokes live.

Pages carrying an ink jcid that the renderer could not place are listed in
`_INK PAGES TO EXPORT.md`. The test is deliberately loose — a page holding typed notes plus a
handwritten annotation counts, and an unrecognised node type that is not a stroke flags the
page anyway.
"""

import argparse
import datetime
import importlib.util
import json
import os
import re
import resource
import shutil
import subprocess
import sys
import traceback
from collections import Counter, defaultdict
from fnmatch import fnmatch
from contextlib import contextmanager
from pathlib import Path


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

GUID = re.compile(r'\((\w{8}-\w{4}-\w{4}-\w{4}-\w{12}), (\d+)\)')
BAD = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


class Encrypted(Exception):
    """The section is password-protected, so there is nothing here to read."""


# Kept out of the body and clearly labelled, because it is machine-read text about the
# attachments rather than anything typed. A downstream reader can split on this line.
OCR_HEADING = (
    '\n\n## Recognised text\n\n'
    "> OneNote's OCR of the attachments above — **not typed**. Kept because the scans and\n"
    '> converted printouts carry no text layer of their own, so this is the only searchable\n'
    '> copy of what they say.\n\n')

# Objects the patched parser refused to read, appended to by patch 5 inside pyOneNote and
# drained per section by parse_pages(). A skipped object is silent data loss, so it is
# counted and reported rather than left to be assumed zero.
SKIPPED = []

# Non-empty when the section is password-protected. See "Password-protected sections".
ENCRYPTED = []

_PATCHES = [
    # 1. ArrayOfPropertyValues (0x10) — unimplemented upstream. MS-ONE 2.6.10: a count, a
    #    PropertyID, then that many PropertySets. `prid` is absent when the count is zero.
    #    The adjacent 0x11 branch calls a five-argument constructor with one argument and
    #    would throw the moment it were reached; fixed alongside.
    ("""            elif type == 0x10:
                raise NotImplementedError('ArrayOfPropertyValues is not implement')
            elif type == 0x11:
                self.rgData.append(PropertySet(file))""",
     """            elif type == 0x10:
                count, = struct.unpack('<I', file.read(4))
                if count:
                    PropertyID(file)
                self.rgData.append([
                    PropertySet(file, OIDs, OSIDs, ContextIDs, document)
                    for _ in range(count)])
            elif type == 0x11:
                self.rgData.append(
                    PropertySet(file, OIDs, OSIDs, ContextIDs, document))"""),

    # 2. Descending into children assumes every node type sets `self.data` with a `.ref`.
    #    Three node types reaching here do not: DataSignatureGroupDefinitionFND has no ref,
    #    and an unrecognised id leaves `data` unset entirely. Skip that subtree instead of
    #    failing the section.
    ("""        current_offset = file.tell()
        if self.file_node_header.baseType == 2:
            self.children.append(FileNodeList(file, self.document, self.data.ref))""",
     """        if self.file_node_header.file_node_type == 'ObjectDataEncryptionKeyV2FNDX':
            _ENCRYPTED.append(1)
        current_offset = file.tell()
        if self.file_node_header.baseType == 2:
            _d = getattr(self, 'data', None)
            if _d is not None and hasattr(_d, 'ref'):
                try:
                    self.children.append(FileNodeList(file, self.document, _d.ref))
                except _BadFragment:
                    pass"""),

    # 11. The hex dump at its source. A four-bytes-of-length property is decoded as UTF-16 and
    #     falls back to `.hex()` on failure — but the failure is routine, not exceptional:
    #     single-byte text of odd length cannot be UTF-16, so short strings are hexed rather
    #     than decoded. Trying UTF-8 first fixes it at the source. Binary still falls through
    #     to hex: latin-1 never raises, so the decision is made on whether the result is ASCII
    #     text, which is the only thing that distinguishes a short string from a short blob.
    ("""                        try:
                            propertyVal = self.rgData[i].Data.decode('utf-16')
                        except:
                            propertyVal = self.rgData[i].Data.hex()""",
     """                        try:
                            propertyVal = self.rgData[i].Data.decode('utf-16')
                        except Exception:
                            _raw = self.rgData[i].Data
                            try:
                                propertyVal = _raw.decode('utf-8')
                            except Exception:
                                _txt = _raw.decode('latin-1')
                                _ascii = sum(32 <= ord(_c) < 127 or _c in '\\r\\n\\t'
                                             for _c in _txt)
                                propertyVal = (_txt if _txt and _ascii >= len(_txt) * 0.9
                                               else _raw.hex())"""),

    # 5. The second unchecked seek. An object declaration jumps to its own `ref` to read a
    #    property set; when that reference is garbage the parse reads noise and dies far away
    #    with an impossible allocation or "rgPrids[i].type is not valid". Bounding the
    #    reference and containing a failed parse costs at most one object — `_SKIPPED` counts
    #    them so the loss is reported rather than assumed to be zero.
    ("""            current_offset = file.tell()
            if self.data.body.jcid.IsPropertySet:
                file.seek(self.data.ref.stp)
                self.propertySet = ObjectSpaceObjectPropSet(file, document)
            file.seek(current_offset)""",
     """            current_offset = file.tell()
            if self.data.body.jcid.IsPropertySet:
                _r = self.data.ref
                _end = _os.fstat(file.fileno()).st_size
                if _r.isFcrNil() or _r.stp + _r.cb > _end:
                    _SKIPPED.append('ref')
                else:
                    file.seek(_r.stp)
                    try:
                        self.propertySet = ObjectSpaceObjectPropSet(file, document)
                    except (ValueError, struct.error, MemoryError, OverflowError,
                            _BadFragment):
                        _SKIPPED.append('parse')
            file.seek(current_offset)"""),

    # 6. The third trusted count, and the one that survived a 2 GB cap. `Count` is 24 bits, so
    #    a desynced read asks for up to 16.7 million CompactIDs. Each is exactly 4 bytes of
    #    file, so a count needing more bytes than remain cannot be real — the bound comes from
    #    the format, not from a guess about plausible sizes.
    ("""class ObjectSpaceObjectStreamOfIDs:
    def __init__(self, file, document):
        self.header = ObjectSpaceObjectStreamHeader(file)
        self.body = []
        self.head = 0
        for i in range(self.header.Count):
            self.body.append(CompactID(file, document))""",
     """class ObjectSpaceObjectStreamOfIDs:
    def __init__(self, file, document):
        self.header = ObjectSpaceObjectStreamHeader(file)
        self.body = []
        self.head = 0
        _room = (_os.fstat(file.fileno()).st_size - file.tell()) // 4
        if self.header.Count > _room:
            _SKIPPED.append('ids')
            self.header.Count = 0
        for i in range(self.header.Count):
            self.body.append(CompactID(file, document))"""),

    # 7. The fourth trusted count, and the one that actually survived every earlier guard.
    #    Bounding the *reference* to a file-data object still leaves `cbLength` — an 8-byte
    #    length read from inside the object — free to demand more memory than the machine has.
    #    The enclosing chunk already declares its own size, so that is the bound.
    ("""        self.guidHeader, self.cbLength, self.unused, self.reserved = struct.unpack('<16sQ4s8s', file.read(36))
        self.FileData, = struct.unpack('{}s'.format(self.cbLength), file.read(self.cbLength))""",
     """        self.guidHeader, self.cbLength, self.unused, self.reserved = struct.unpack('<16sQ4s8s', file.read(36))
        _room = min(fileNodeChunkReference.cb,
                    _os.fstat(file.fileno()).st_size - file.tell())
        if self.cbLength > _room:
            _SKIPPED.append('filedata')
            self.cbLength = 0
        self.FileData, = struct.unpack('{}s'.format(self.cbLength), file.read(self.cbLength))"""),

    # 8. Rendering a CompactID looks its index up in the revision's id table and assumes it is
    #    there. A garbage index raises KeyError from __str__, which kills the whole section for
    #    one unreadable identity. Degrading to a marker is safe rather than lossy-by-stealth:
    #    parse_pages matches identities by GUID regex, so an unresolved one fails to match and
    #    its property is dropped, never misattributed to the wrong page.
    ("""    def __str__(self):
        return '<ExtendedGUID> ({}, {})'.format(
        self.document._global_identification_table[self.current_revision][self.guidIndex],
        self.n)""",
     """    def __str__(self):
        try:
            _guid = self.document._global_identification_table[
                self.current_revision][self.guidIndex]
        except KeyError:
            _SKIPPED.append('identity')
            return '<ExtendedGUID> (unresolved, {})'.format(self.n)
        return '<ExtendedGUID> ({}, {})'.format(_guid, self.n)"""),

    # 9. The last trusted length. `cb` is four bytes straight from the file, so a desynced read
    #    asks for up to 4 GB — and the value is later rendered with `.hex()`, which doubles it
    #    again. Nothing can be longer than the file that contains it.
    ("""class PrtFourBytesOfLengthFollowedByData:
    def __init__(self, file, propertySet):
        self.cb, = struct.unpack('<I', file.read(4))
        self.Data, = struct.unpack('{}s'.format(self.cb), file.read(self.cb))""",
     """class PrtFourBytesOfLengthFollowedByData:
    def __init__(self, file, propertySet):
        self.cb, = struct.unpack('<I', file.read(4))
        _room = _os.fstat(file.fileno()).st_size - file.tell()
        if self.cb > _room:
            _SKIPPED.append('length')
            self.cb = 0
        self.Data, = struct.unpack('{}s'.format(self.cb), file.read(self.cb))"""),

    # 10. The format's own integrity check, unused upstream. MS-ONESTORE 2.4.2 requires every
    #    FileNodeListFragment to open with uintMagic == 0xA4567AB1F5F7F4C4, and requires a
    #    reference to lie inside the file. Neither is verified, so one garbage reference sends
    #    the traversal into arbitrary bytes and every structure read after it is noise —
    #    surfacing much later as an impossible allocation or a read past EOF, with nothing to
    #    say where it began. Checking the magic turns that into a bounded, detectable loss.
    ("""class FileNodeList:
    def __init__(self, file, document, file_chunk_reference):
        file.seek(file_chunk_reference.stp)""",
     """class _BadFragment(Exception):
    \"\"\"A reference did not lead to a valid FileNodeListFragment.\"\"\"


class FileNodeList:
    def __init__(self, file, document, file_chunk_reference):
        _end = _os.fstat(file.fileno()).st_size
        if (file_chunk_reference.isFcrNil() or file_chunk_reference.cb <= 0
                or file_chunk_reference.stp + file_chunk_reference.cb > _end):
            raise _BadFragment('reference outside the file')
        file.seek(file_chunk_reference.stp)"""),

    ("""class FileNodeListHeader:
    def __init__(self, file):
        self.uintMagic, self.FileNodeListID, self.nFragmentSequence = struct.unpack('<8sII', file.read(16))""",
     """class FileNodeListHeader:
    MAGIC = b'\\xc4\\xf4\\xf7\\xf5\\xb1\\x7aV\\xa4'

    def __init__(self, file):
        raw = file.read(16)
        if len(raw) < 16:
            raise _BadFragment('fragment header past end of file')
        self.uintMagic, self.FileNodeListID, self.nFragmentSequence = struct.unpack('<8sII', raw)
        if self.uintMagic != FileNodeListHeader.MAGIC:
            raise _BadFragment('fragment magic mismatch')"""),

    # 3. A file-data reference is followed without checking it. `isFcrNil()` is defined
    #    upstream and never called anywhere, and nothing bounds the offset against the file,
    #    so a nil or out-of-range reference reads past EOF and raises struct.error.
    ("""        current_offset = file.tell()
        file.seek(self.ref.stp)
        self.fileDataStoreObject = FileDataStoreObject(file, self.ref)
        file.seek(current_offset)""",
     """        current_offset = file.tell()
        _end = _os.fstat(file.fileno()).st_size
        if self.ref.isFcrNil() or self.ref.stp + self.ref.cb > _end:
            self.fileDataStoreObject = None
        else:
            file.seek(self.ref.stp)
            self.fileDataStoreObject = FileDataStoreObject(file, self.ref)
        file.seek(current_offset)"""),

    # 4. `count` here is read straight from the file and used as a loop bound, so a garbage
    #    value builds a list of billions of entries and exhausts the machine. The stream it
    #    reads from holds a known number of ids; asking for more than exist is proof the
    #    value is garbage, not a bigger array.
    ("""    @staticmethod
    def get_compact_ids(stream_of_context_ids, count):
        data = []
        for i in range(count):
            data.append(stream_of_context_ids.read())
        return data""",
     """    @staticmethod
    def get_compact_ids(stream_of_context_ids, count):
        count = min(count, len(stream_of_context_ids.body))
        data = []
        for i in range(count):
            data.append(stream_of_context_ids.read())
        return data"""),
]

# Patch 3 can legitimately leave `fileDataStoreObject` unset, and the collector that reads it
# assumes it is always there — so refusing a bad reference has to be handled at both ends.
_DOC_PATCHES = [
    ("""                    self._files[str(node.data.guidReference)]["content"] = node.data.fileDataStoreObject.FileData""",
     """                    if node.data.fileDataStoreObject is not None:
                        self._files[str(node.data.guidReference)]["content"] = node.data.fileDataStoreObject.FileData"""),

    # One blob is declared once per revision that referenced it, each declaration carrying a
    # different identity, and `_files[guid]["identity"] = ...` keeps only the last. Every page
    # pointing at an earlier identity then resolves to nothing. Accumulating them costs a list
    # per blob; see "Why an image needs every identity" in the module docstring.
    ("""                    self._files[guid]["identity"] = str(node.data.oid)""",
     """                    self._files[guid]["identity"] = str(node.data.oid)
                    self._files[guid].setdefault("identities", []).append(str(node.data.oid))"""),
]


def _patch_module(name, patches, extra=None):
    """Import `name` from pyOneNote with `patches` applied to its source, in memory.

    Patched here rather than in site-packages so the fixes travel with this file — a
    hand-edited virtualenv is not a fix anyone else can reproduce. Every patch anchors on an
    exact source excerpt and is fatal if it stops matching, so an upgrade cannot silently
    leave the parser unguarded.
    """
    spec = importlib.util.find_spec(name)
    src = Path(spec.origin).read_text()
    for i, (old, new) in enumerate(patches, 1):
        if old not in src:
            raise SystemExit(f'pyOneNote has changed — {name} patch {i} no longer applies')
        src = src.replace(old, new)
    module = importlib.util.module_from_spec(spec)
    module.__dict__.update(extra or {})
    sys.modules[name] = module
    exec(compile(src, spec.origin, 'exec'), module.__dict__)
    return module


def _patched_pyonenote():
    """Load pyOneNote with every patch applied, plus one behavioural change made here:
    unrecognised node types are all named 'Unknown' upstream, which makes ink
    indistinguishable from any other unsupported structure. Keeping the jcid gives something
    to correlate.
    """
    module = _patch_module('pyOneNote.FileNode', _PATCHES,
                           {'_os': os, '_SKIPPED': SKIPPED, '_ENCRYPTED': ENCRYPTED})

    for name in dir(module):
        cls = getattr(module, name)
        mapping = getattr(cls, '_jcid_name_mapping', None)
        if isinstance(mapping, dict):
            def _get(self, _m=mapping):
                return _m.get(self.jcid, f'Unknown(0x{self.jcid:08x})')
            for attr in ('get_jcid_name', 'jcid_name'):
                if hasattr(cls, attr):
                    setattr(cls, attr, _get)

    _patch_module('pyOneNote.OneDocument', _DOC_PATCHES)
    from pyOneNote.OneDocument import OneDocment
    return OneDocment


_HEX = re.compile(r'^(?:[0-9a-f]{2})+$')


def _decode(text: str, prop_name: str) -> str:
    """Recover text from the two ways pyOneNote mangles it.

    1. `TextExtendedAscii` holds single-byte text, but the library decodes every string as
       UTF-16, so byte pairs come through as CJK code points. Re-encoding to UTF-16LE
       recovers the bytes exactly.
    2. Values carried by the four-bytes-of-length-followed-by-data property type arrive as a
       hex dump: `4a616e7561727920...` decodes to "January ...".

    Case 2 is fixed upstream by patch 11, which tries UTF-8 before falling back to hex, so
    this branch is a safety net rather than the main path — a length guard cannot be the main
    path, because the shortest casualties are the commonest.

    The guard stays because a recovery pass must not destroy anything: short real text can be
    all hex characters ("added", "facade"), so it wants an even length of at least 12 and an
    almost entirely printable result. A missed decode leaves hex visible; the reverse silently
    destroys real text.
    """
    if prop_name == 'TextExtendedAscii':
        text = text.encode('utf-16-le').decode('ascii', 'replace').replace('\x00', '')
    stripped = text.strip().lower()
    if len(stripped) >= 12 and len(stripped) % 2 == 0 and _HEX.match(stripped):
        try:
            candidate = bytes.fromhex(stripped).decode('latin-1')
        except ValueError:
            return text
        printable = sum(c.isprintable() or c in '\r\n\t' for c in candidate)
        if printable >= len(candidate) * 0.9:
            return candidate
    return text


def _clean(text: str) -> str:
    text = text.replace(chr(0), '').replace('\r', '\n')
    text = re.sub('[﻿   -]', ' ', text)
    text = re.sub(r'HYPERLINK "[^"]*"', '', text)
    return re.sub(r'[ \t]{2,}', ' ', text).strip()


def _safe(name: str, fallback: str) -> str:
    """A filename Windows, Nextcloud and Obsidian will all accept.

    **The length cap applies to the stem, never the extension.** Callers pass whole filenames
    as well as page titles, and cutting `... Community 01.png` at 80 characters left
    `... Community 01.` — a trailing dot, which Nextcloud rejects outright, Windows clients
    refuse to sync, and which loses the file's type. Stripping dots before the cut does not
    help: the cut is what creates the last one, so it has to run again afterwards.

    A suffix only counts as an extension when it is alphanumeric and at most six characters,
    so a page title like `Meeting notes 3.5 hours` is left alone.
    """
    clean = BAD.sub('-', name).strip().strip('.') or fallback
    stem, dot, ext = clean.rpartition('.')
    if not (dot and 0 < len(ext) <= 6 and ext.isalnum()):
        stem, dot, ext = clean, '', ''
    return (stem[:80 - len(dot + ext)].strip().strip('.') or fallback[:80]) + dot + ext


def parse_pages(path: Path, OneDocment):
    """-> (pages, files_by_identity, note). Raises if the file cannot be parsed.

    Raises `Encrypted` for a password-protected section rather than returning nothing, so it
    cannot be mistaken for an empty one — see "Password-protected sections".
    """
    SKIPPED.clear()
    ENCRYPTED.clear()
    with open(path, 'rb') as f:
        doc = OneDocment(f)
        # get_properties(), never get_json(): the latter hex-encodes every embedded file
        # into a string, doubling the bytes on top of the originals it keeps, to build a
        # dict of which only `properties` is used. Large sections exhaust memory on it.
        properties = doc.get_properties()
        files = doc.get_files()

    # Checked after the parse, not before: the marker nodes are only seen while walking.
    if ENCRYPTED and not properties:
        raise Encrypted(f'password-protected ({len(ENCRYPTED)} encryption-key nodes)')

    spaces = defaultdict(list)
    for prop in properties:
        m = GUID.search(str(prop.get('identity', '')))
        if m:
            spaces[m.group(1)].append(prop)

    pages = []
    for guid, props in spaces.items():
        if not any(p['type'] == 'jcidPageNode' for p in props):
            continue                        # section metadata, version history, etc.
        # `space` is the object-space GUID. It changes nothing here, but it is the only
        # stable handle onto a page, and onenote2md-ink needs it to put a rendered
        # stroke on the right page without re-deriving the ordering.
        page = {'title': '', 'created': '', 'space': guid, 'text': [], 'ocr': [],
                'images': [], 'unknown': Counter()}
        for p in props:
            val = p.get('val') or {}
            blob = json.dumps(val, default=str)
            if p['type'] == 'jcidPageMetaData':
                page['title'] = page['title'] or _clean(str(val.get('CachedTitleString', '')))
                page['created'] = page['created'] or str(val.get('TopologyCreationTimeStamp', ''))
            if p['type'] in ('jcidImageNode', 'jcidEmbeddedFileNode'):
                for key in ('PictureContainer', 'EmbeddedFileContainer'):
                    for ref in (val.get(key) or []):
                        page['images'].append(
                            (str(ref), _clean(str(val.get('ImageAltText', '')))))
            if p['type'].startswith('Unknown'):
                page['unknown'][p['type']] += 1
            # Which object carried the text decides what it *is*. Text on an image or an
            # embedded file is OneNote's OCR of that attachment, not something typed. They
            # are separated here because they cannot be separated afterwards.
            bucket = 'ocr' if p['type'] in ('jcidImageNode', 'jcidEmbeddedFileNode') else 'text'
            for prop_name in ('RichEditTextUnicode', 'TextExtendedAscii'):
                for raw in re.findall(f'"{prop_name}": "(.*?)(?<!\\\\)"', blob):
                    try:
                        t = _clean(_decode(json.loads(f'"{raw}"'), prop_name))
                    except Exception:
                        continue
                    if (t and t != page['title']
                            and t not in page['text'] and t not in page['ocr']):
                        page[bucket].append(t)
        # One image node appears once per revision that touched it, and each occurrence
        # contributes its PictureContainer again — so a two-page scan was written as four
        # files, and a merged PDF would silently contain every page twice. Text was already
        # deduped on the way in; images never were. Dedupe on the **reference**, not on the
        # bytes: two distinct references to identical bytes are an image genuinely placed
        # twice on the page, and that is not ours to collapse.
        seen, unique = set(), []
        for ref, alt in page['images']:
            if ref in seen:
                continue
            seen.add(ref)
            unique.append((ref, alt))
        page['images'] = unique
        pages.append(page)

    pages.sort(key=lambda p: (p['created'] or '9999', p['title']))

    # Every identity a blob was declared under, not just the last one to be written.
    by_identity, blob_of = {}, {}
    for guid, v in files.items():
        for ident in (v.get('identities') or ([v['identity']] if v.get('identity') else [])):
            by_identity.setdefault(ident, v)
            blob_of.setdefault(ident, guid)

    wanted = [ref for p in pages for ref, _ in p['images']]
    unresolved = [r for r in wanted if r not in by_identity]
    # Resolving is not the same as having bytes: a bounded-away reference leaves an entry with
    # no content, which write_section skips. Counting those as linked overstated what was
    # written by 27 images across the recovered sections.
    empty = [r for r in wanted if r in by_identity and not by_identity[r].get('content')]
    orphans = len(files) - len({blob_of[r] for r in wanted if r in blob_of})
    note = (f'{len(wanted) - len(unresolved) - len(empty)}/{len(wanted)} images linked'
            + (f', {len(unresolved)} UNRESOLVED' if unresolved else '')
            + (f', {len(empty)} EMPTY' if empty else '')
            + (f', {orphans} unused (old revisions)' if orphans else '')
            + (', SKIPPED ' + ' '.join(f'{n}x{k}' for k, n in
                                       sorted(Counter(SKIPPED).items())) if SKIPPED else ''))
    return pages, by_identity, note


def is_ink(page) -> bool:
    """Any node type the parser cannot read — almost always stylus handwriting.

    A false positive costs one unnecessary export, so the test is deliberately loose.
    """
    return bool(page['unknown'])


def ink_confidence(page) -> str:
    if sum(len(t) for t in page['text']) < 40 and not page['images']:
        return 'high — nothing else on the page'
    return 'mixed — page also has typed text or images'


def _fm(value: str) -> str:
    """Flatten a value for a frontmatter field. A page title may contain a newline."""
    return re.sub(r'\s+', ' ', str(value)).strip()


def _to_pdf(data: bytes) -> bytes:
    """An XPS printout as PDF bytes. Raises if it cannot be converted.

    "Print to OneNote" output is stored as XPS, which little outside Windows opens. Nothing is
    lost by writing the PDF instead: the XPS draws its glyphs as outlines rather than text, so
    neither form has a text layer and neither is searchable without OCR. The PDF is simply the
    one that opens. The `.one` sources remain the originals either way.
    """
    import pymupdf
    doc = pymupdf.open(stream=data, filetype='xps')
    try:
        return doc.convert_to_pdf()
    finally:
        doc.close()


def _convert_one(path: str):
    """Worker: one XPS → a sibling PDF. Returns (path, error or None).

    Module level so it pickles for the pool, and it sets its own umask because that is
    per-process — a worker inheriting the default ignores `--umask`.
    """
    xps = Path(path)
    if SHARE_UMASK is not None:
        os.umask(SHARE_UMASK)
    try:
        xps.with_suffix('.pdf').write_bytes(_to_pdf(xps.read_bytes()))
    except Exception as e:
        return path, f'{type(e).__name__}: {str(e)[:60]}'
    return path, None


def _repoint(converted, log):
    """Point each page's `![[...]]` link at the PDF instead of the XPS.

    Done in the parent and grouped per markdown file. Several `.xps` can share one page, so
    their links live in the same `.md`; rewriting that file once per attachment from parallel
    workers is a read-modify-write race that silently drops a link.

    Run over every converted file, not just this run's, so a run interrupted before repointing
    is repaired by the next one rather than leaving links pointing at an XPS forever.
    """
    by_md = defaultdict(list)
    for xps in converted:
        by_md[xps.parent.parent / (xps.parent.name[:-len('_files')] + '.md')].append(xps)

    n = 0
    for md, items in by_md.items():
        if not md.exists():
            continue
        text = original = md.read_text(errors='replace')
        for xps in items:
            text = text.replace(f'{xps.parent.name}/{xps.name}',
                                f'{xps.parent.name}/{xps.with_suffix(".pdf").name}')
        if text != original:
            md.write_text(text)
            n += 1
    log(f'  repointed {n} pages')


def convert_xps(out: Path, log, workers: int = 5) -> tuple:
    """Convert `.xps` already written under `out`, and repoint the markdown that links them.

    **The only place XPS is converted.** A run calls this after its sections; `--convert-xps`
    calls it alone, for output extracted before it existed. `write_section` used to convert
    inline as well, one file at a time on the main thread, so the same job existed twice and
    the copy every normal extraction used was the slow one — a fifth of the throughput, on the
    kind of section that holds a thousand multi-page printouts.

    Idempotent twice over: a file whose PDF exists is skipped, and repointing re-runs over
    everything, so an interrupted run costs only what it had not reached. That is what makes it
    safe to call unconditionally at the end of a run.

    The work is CPU-bound in MuPDF, not I/O-bound, so it runs in a process pool.

    Originals are kept — the source export holds them anyway, but deleting files is the
    caller's decision, so the reclaimable size is reported instead.
    """
    from concurrent.futures import ProcessPoolExecutor

    found = sorted(out.rglob('*.xps'))
    if not found:
        log('no .xps found — nothing to convert')
        return 0, 0, 0

    done = [x for x in found if x.with_suffix('.pdf').exists()]
    todo = [x for x in found if not x.with_suffix('.pdf').exists()]
    log(f'{len(found)} .xps — {len(done)} already converted, {len(todo)} to do '
        f'on {workers} workers')

    ok, bad = list(done), []
    if todo:
        with ProcessPoolExecutor(max_workers=workers) as pool:
            for i, (path, err) in enumerate(
                    pool.map(_convert_one, [str(x) for x in todo], chunksize=4), 1):
                if err:
                    bad.append(path)
                    log(f'  FAILED {Path(path).relative_to(out)}  {err}')
                else:
                    ok.append(Path(path))
                if i % 200 == 0:
                    log(f'  ...{i}/{len(todo)}')

    with share_umask():
        _repoint(ok, log)

    reclaimable = sum(x.stat().st_size for x in ok if x.exists())
    log(f'\nXPS  {len(ok)} converted, {len(bad)} failed, '
        f'{reclaimable / 1e9:.2f} GB of .xps now redundant (kept — delete when satisfied)')
    return len(ok), len(bad), reclaimable


def write_section(notebook, section, path, out_dir, OneDocment, log, dead, resume=False):
    sec_dir = out_dir / _safe(notebook, 'Notebook') / _safe(section, 'Section')

    # A section's directory is created only after its parse succeeds, so "output exists"
    # means "this one worked". Failures therefore retry on every run, which is what makes
    # --resume safe: it can skip nothing that still needs doing.
    if resume and sec_dir.is_dir() and any(sec_dir.glob('*.md')):
        log(f'  {section:<42} skipped — already extracted')
        return 0, 0, 0

    try:
        pages, files, note = parse_pages(path, OneDocment)
    except Encrypted as e:
        # Distinct from unparseable, because the fix is manual and belongs to the notebook owner:
        # is not enough, the password has to be *removed* and the notebook re-exported.
        log(f'  {section:<42} ENCRYPTED    {e}')
        dead.append((notebook, section, f'password-protected — remove the password in '
                                        f'OneNote (unlocking is not enough) and re-export'))
        return 0, 0, 0
    except Exception as e:
        log(f'  {section:<42} UNPARSEABLE  {type(e).__name__}: {str(e)[:60]}')
        dead.append((notebook, section, f'{type(e).__name__}: {e}'))
        return 0, 0, 0

    n_ink = n_img = n_xps = 0
    with share_umask():
        sec_dir.mkdir(parents=True, exist_ok=True)
        for i, page in enumerate(pages, 1):
            title = _safe(page['title'], f'Untitled page {i}')
            ink_page = is_ink(page)
            n_ink += ink_page
            links = []
            mine = [(r, a) for r, a in page['images'] if r in files]
            if mine:
                fdir = sec_dir / f'{i:03d} {title}_files'
                fdir.mkdir(exist_ok=True)
                for k, (ref, alt) in enumerate(mine, 1):
                    data = files[ref].get('content')
                    if not data:
                        continue
                    ext = (files[ref].get('extension') or '').lstrip('.').lower() or 'bin'
                    # Written as `.xps` and converted afterwards by `convert_xps`, in a pool.
                    # Converting here meant doing it one file at a time on the main thread
                    # while the other cores idled. That pass already existed, is already
                    # idempotent and already repoints the links, so this is one implementation
                    # rather than two that have drifted.
                    n_xps += ext == 'xps'
                    fname = f'{k:02d}.{ext}'
                    (fdir / fname).write_bytes(data)
                    links.append(f'![[{fdir.name}/{fname}]]'
                                 + (f'\n*{alt}*' if alt and len(alt) < 120 else ''))
                    n_img += 1
            (sec_dir / f'{i:03d} {title}.md').write_text(
                '---\n'
                f'title: {_fm(page["title"]) or f"Untitled page {i}"}\n'
                f'notebook: {_fm(notebook)}\nsection: {_fm(section)}\n'
                f'created: {page["created"][:10]}\nsource: onenote\n'
                f'images: {len(links)}\n'
                f'typed_chars: {sum(len(t) for t in page["text"])}\n'
                f'ocr_chars: {sum(len(t) for t in page["ocr"])}\n'
                + ('handwriting: true\n'
                   f'ink_confidence: {ink_confidence(page)}\n'
                   f'ink_nodes: {", ".join(sorted(page["unknown"]))}\n' if ink_page else '')
                + f'note: "{note}"\n---\n\n'
                f'# {page["title"] or f"Untitled page {i}"}\n\n'
                + ('> **Content here cannot be extracted** — most likely stylus handwriting,\n'
                   '> which OneNote stores as vector ink rather than pixels. Export this page\n'
                   '> from OneNote as PDF. See `_INK PAGES TO EXPORT.md`.\n\n' if ink_page else '')
                + ('\n\n'.join(page['text']) or '_(nothing typed on this page)_')
                + ('\n\n' + '\n'.join(links) if links else '')
                + (OCR_HEADING + '\n\n'.join(page['ocr']) if page['ocr'] else '') + '\n')

    log(f'  {section:<42} {len(pages):>4} pages  {n_img:>4} images  {n_ink:>3} ink'
        + (f'  {n_xps:>4} xps' if n_xps else '')
        + f'  [{note}]')
    return len(pages), n_img, n_ink


def _sections(pkg: Path, work: Path, only=None):
    """Yield (section name, extracted path) for each .one inside a .onepkg (a CAB).

    **Unpacked with `cabextract`, not libarchive, and that is not a preference.** A `.onepkg`
    is a CAB compressed with LZX, and libarchive before 3.8.9 returns the right number of
    bytes with the wrong values in them — silently. Both decoders agree on the length of
    `Filing/Other.one` (119,014,976) and differ in 276,291 bytes; carving CRC-checked PNGs
    out of each tells the story:

        libarchive   15 valid PNGs, 68 broken
        cabextract   83 valid PNGs,  0 broken

    **The trigger is position.** LZX may rewrite the four-byte operand after each `0xE8`
    (x86 `CALL`) opcode and the decoder must undo it, but the spec stops translating after
    the first 1 GiB of a folder's output. libarchive had no such limit, so every coincidental
    `0xE8` inside already-compressed data past that mark had the following four bytes
    rewritten. `Insurances.one` proves it by straddling the boundary: it starts at folder
    offset 977,167,482, 1 GiB falls 96,574,342 bytes in, and the first differing byte is at
    96,593,691 with nothing before it differing at all.

    So the damage is computable, not mysterious: sum the CAB directory's uncompressed sizes
    and any section extending past 1 GiB was corrupted. In one 1.3 GB notebook that was four
    sections — 112 undecodable images, and one that failed to parse at all.

    libarchive fixed this in **3.8.9** (commit `220eefe34`, 2026-07-28), so requiring that
    version is a defensible alternative. It is not the one taken here: `libarchive-c` binds
    whatever system library is present and cannot pin a version at install time, so enforcing
    it means a runtime check and a hard error — against a 200 KB apt package and a
    `shutil.which()`. Distributions still shipping an affected version include Ubuntu 24.04
    (3.7.2) and Debian 12 (3.6.2).
    """
    if not shutil.which('cabextract'):
        raise SystemExit(
            'cabextract is required to unpack .onepkg — `sudo apt install cabextract` '
            '(or `brew install cabextract`).\n'
            'libarchive is not a fallback: before 3.8.9 its LZX decoder silently corrupts '
            'any part of a cabinet past the first 1 GiB.')
    dest = work / 'sections'
    shutil.rmtree(dest, ignore_errors=True)
    dest.mkdir(parents=True, exist_ok=True)
    # `only` pulls just the named sections out of the cabinet. A `.onepkg` can be well over a
    # gigabyte of `.one`, and a caller wanting two of them should not pay to unpack nineteen.
    patterns = [f'*{s}.one' for s in only] if only else ['*.one']
    cmd = ['cabextract', '-q']
    for pat in patterns:
        cmd += ['-F', pat]
    r = subprocess.run(cmd + ['-d', str(dest), str(pkg)], capture_output=True, text=True)
    if r.returncode != 0 and not any(dest.rglob('*.one')):
        raise RuntimeError(f'cabextract failed on {pkg.name}: '
                           f'{(r.stderr or r.stdout).strip()[:200]}')
    for one in sorted(dest.rglob('*.one')):
        name = str(one.relative_to(dest)).replace('\\', '/')
        yield re.sub(r'\.one$', '', name, flags=re.I).replace('/', ' - '), one


def _ink_rows(out: Path) -> list:
    """Ink pages read back from the written markdown, not collected during the run.

    A resumed run re-extracts only what is missing, so a checklist built from the run itself
    would drop every ink page an earlier run had already found — quietly, and in the one file
    whose entire job is to be the complete list.
    """
    rows = []
    for md in sorted(out.rglob('*.md')):
        if md.name.startswith('_'):
            continue
        head = md.read_text(errors='replace').split('\n---\n', 1)[0]
        meta = dict((k.strip(), v.strip())
                    for k, _, v in (line.partition(':') for line in head.splitlines()))
        if meta.get('handwriting') == 'true':
            rows.append(tuple(meta.get(k, '') for k in (
                'notebook', 'section', 'title', 'created', 'ink_confidence', 'ink_nodes')))
    return sorted(rows)


def _now() -> str:
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M')


def _page_meta(out: Path):
    """Every extracted page's frontmatter, from disk. The one reader of the written output."""
    for md in sorted(out.rglob('*.md')):
        if md.name.startswith('_'):
            continue
        head = md.read_text(errors='replace').split('\n---\n', 1)[0]
        yield dict((k.strip(), v.strip())
                   for k, _, v in (line.partition(':') for line in head.splitlines()))


def _corpus_summary(out: Path) -> str:
    """Per-section figures rebuilt by reading the written markdown back.

    Written from the run's own log lines instead, a `--resume` run replaces the whole file with
    the handful of sections it did not skip and every earlier section's figures are gone. That
    is the trap `_ink_rows()` already existed to avoid, in a file with the same job: describe
    the corpus, not the run. Everything needed is in the frontmatter.
    """
    sections, order = {}, []
    for meta in _page_meta(out):
        key = (meta.get('notebook', ''), meta.get('section', ''))
        if key not in sections:
            sections[key] = {'pages': 0, 'images': 0, 'ink': 0, 'note': meta.get('note', '')}
            order.append(key)
        row = sections[key]
        row['pages'] += 1
        row['images'] += int(meta.get('images', 0) or 0)
        row['ink'] += meta.get('handwriting') == 'true'
        # `note` is a property of the section's parse, so any page carries it; keep the first
        # non-empty one rather than the last, which is blank on most pages.
        if not row['note'].strip('"'):
            row['note'] = meta.get('note', '')

    lines = [f'# The corpus on disk, as of {_now()}', '',
             'Rebuilt by reading every extracted page back, so it describes the whole extract',
             'and not just the sections the last run touched. Per-run transcripts, including',
             'failures and skips, are in `_extraction runs.txt`.', '']
    total = [0, 0, 0]
    for notebook, section in order:
        r = sections[(notebook, section)]
        total = [a + b for a, b in zip(total, (r['pages'], r['images'], r['ink']))]
        note = r['note'].strip('"')
        lines.append(f'{notebook} / {section:<42} {r["pages"]:>4} pages  {r["images"]:>5} images  '
                     f'{r["ink"]:>3} ink' + (f'  [{note}]' if note else ''))
    lines += ['', f'{len(order)} sections  {total[0]} pages  {total[1]} images  {total[2]} ink '
                  f'pages', '']
    return '\n'.join(lines)


def _checklist(ink, dead) -> str:
    lines = [
        '# To export from OneNote by hand', '',
        'OneNote saves stylus input as vector ink strokes rather than pixels, so there is',
        'no image to recover — these pages carry node types the parser cannot read at all.',
        'Open each in OneNote and use Print / Export as PDF.',
        '',
        '*high* confidence means nothing else came through, so the page is probably',
        'handwriting and nothing else. *mixed* means text or images extracted normally and',
        'the ink sits alongside them — the extracted note is worth reading but incomplete.',
        '', '| Notebook | Section | Page | Created | Confidence | Unreadable node types |',
        '|---|---|---|---|---|---|',
    ]
    lines += ['| ' + ' | '.join(x or '' for x in row) + ' |' for row in sorted(ink)]
    lines += ['', f'{len(ink)} pages.', '', '## Whole sections that could not be read', '',
              'Nothing was extracted from these. Two causes, and they need different fixes:',
              '',
              '- **password-protected** — open the section in OneNote and use Password',
              '  Protection → **Remove Password**, then re-export. *Unlocking is not enough*:',
              '  it decrypts the section for that session only, and an export still carries',
              '  the encrypted bytes.',
              '- **undocumented `638de92f` format** — re-export the notebook as `.onepkg`,',
              '  which writes it in the documented one.',
              '',
              'Either way, re-run with `--resume`; a section with no output is retried.', '',
              '| Notebook | Section | Why |', '|---|---|---|']
    lines += [f'| {nb} | {sec} | {why[:70]} |' for nb, sec, why in sorted(dead)]
    return '\n'.join(lines + ['', f'{len(dead)} sections.', ''])


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('--source', help='folder searched for *.onepkg and *.one')
    ap.add_argument('--out', required=True, help='where to write the markdown')
    ap.add_argument('--work', default='/tmp/onenote2md', help='scratch for unpacking')
    ap.add_argument('--max-memory-gb', type=float, default=2.0,
                    help='address-space cap; see "Why the memory cap exists"')
    ap.add_argument('--resume', action='store_true',
                    help='skip sections already extracted; retries failures and does the rest')
    ap.add_argument('--convert-xps', action='store_true',
                    help='convert .xps already under --out to PDF; no extraction, no --source')
    ap.add_argument('--workers', type=int, default=5,
                    help='parallel processes for XPS conversion (default 5)')
    ap.add_argument('--exclude', action='append', default=[], metavar='GLOB',
                    help='skip source files matching this (repeatable). For a notebook that '
                         'has been superseded by a re-export — leaving both in place '
                         'extracts each of them, under two notebook names')
    ap.add_argument('--umask', type=lambda v: int(v, 8), metavar='OCTAL',
                    help='umask for files written, e.g. 007 when a group needs write access '
                         'on a network share. Default: leave the process umask alone')
    args = ap.parse_args()
    if args.umask is not None:
        globals()['SHARE_UMASK'] = args.umask
    if not args.source and not args.convert_xps:
        ap.error('--source is required unless --convert-xps')

    cap = int(args.max_memory_gb * (1 << 30))
    resource.setrlimit(resource.RLIMIT_AS, (cap, resource.getrlimit(resource.RLIMIT_AS)[1]))

    out, work = Path(args.out), Path(args.work)
    log_lines, dead = [], []

    def log(m):
        print(m, flush=True)
        log_lines.append(m)

    if args.convert_xps:
        convert_xps(out, log, args.workers)
        return

    try:
        OneDocment = _patched_pyonenote()
        import pymupdf  # noqa: F401
    except ImportError as e:
        raise SystemExit(f'{e}\n\n    pip install pyOneNote pymupdf')

    src = Path(args.source)
    work.mkdir(parents=True, exist_ok=True)

    def wanted(path):
        keep = not any(fnmatch(path.name, g) for g in args.exclude)
        if not keep:
            log(f'skipping {path.name} (--exclude)')
        return keep

    totals = [0, 0, 0]
    for pkg in sorted(p for p in src.rglob('*.onepkg') if wanted(p)):
        log(f'\n### {pkg.stem}')
        for section, tmp in _sections(pkg, work):
            try:
                got = write_section(pkg.stem, section, tmp, out, OneDocment, log, dead,
                                    args.resume)
                totals = [a + b for a, b in zip(totals, got)]
            except Exception:
                log(f'  {section}: FAILED\n{traceback.format_exc()[-300:]}')

    loose = sorted(p for p in src.rglob('*.one') if wanted(p))
    if loose:
        log('\n### loose sections')
        for one in loose:
            got = write_section(one.parent.name, one.stem, one, out, OneDocment, log, dead,
                                args.resume)
            totals = [a + b for a, b in zip(totals, got)]

    log(f'\nTOTAL  {totals[0]} pages  {totals[1]} images  {totals[2]} ink pages')

    # The sections wrote `.xps` untouched; convert them all here, in the pool. Doing it as part
    # of the run rather than leaving it to a second command keeps one invocation doing the whole
    # job — and the pass is idempotent, so it costs nothing when there is none.
    if any(out.rglob('*.xps')):
        log('')
        convert_xps(out, log, args.workers)

    with share_umask():
        out.mkdir(parents=True, exist_ok=True)
        (out / '_INK PAGES TO EXPORT.md').write_text(_checklist(_ink_rows(out), dead))
        # Two files, each complete by construction. The summary is rebuilt from disk so a
        # resumed run cannot leave it describing only the sections it touched; the transcript
        # is appended so no run's failures are ever written over.
        (out / '_extraction log.txt').write_text(_corpus_summary(out))
        with (out / '_extraction runs.txt').open('a') as f:
            f.write(f'\n{"=" * 78}\n=== run {_now()}\n{"=" * 78}\n'
                    + '\n'.join(log_lines) + '\n')


if __name__ == '__main__':
    main()
