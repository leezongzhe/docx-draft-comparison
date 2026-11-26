from __future__ import annotations

import math
import os
import re
import shutil
import subprocess
import tempfile
from bisect import bisect_left
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Callable, Iterable, Sequence

import pdfplumber
import pypdfium2 as pdfium
from PIL import Image, ImageDraw
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A3, landscape
from reportlab.pdfgen import canvas


Progress = Callable[[str], None]
# Keep complete decimal values as one token.  The previous expression stopped at
# the decimal point (for example, both ``0.50`` and ``0.75`` became ``0``), so
# repeated values in document tables could be aligned to the wrong row and a
# genuinely changed cell could be left unmarked.
TOKEN_RE = re.compile(
    r"\d+(?:[.,]\d+)*|[^\W\d_]+(?:['’-][^\W\d_]+)*",
    re.UNICODE,
)


class ComparisonError(RuntimeError):
    pass


@dataclass
class Word:
    text: str
    x0: float
    top: float
    x1: float
    bottom: float
    page: int
    fontname: str = ""
    size: float = 0.0

    @property
    def token(self) -> str:
        match = TOKEN_RE.search(self.text.casefold())
        # Decorative rules, chart bars, bullets, and equation-only glyphs do not
        # provide reliable word correspondence. Preserve them visually, but do
        # not turn renderer glyphs into false textual additions or deletions.
        return match.group(0) if match else ""


@dataclass
class Block:
    page: int
    index: int
    words: list[Word]
    text: str
    global_index: int = 0

    @property
    def tokens(self) -> list[str]:
        return [word.token for word in self.words if word.token]

    @property
    def normalized(self) -> str:
        return " ".join(self.tokens)


@dataclass
class PageData:
    number: int
    width: float
    height: float
    words: list[Word]
    blocks: list[Block]


@dataclass
class BlockMatch:
    left: Block
    right: Block
    score: float
    moved: bool = False


@dataclass
class Mark:
    page: int
    rect: tuple[float, float, float, float]
    kind: str
    outline_only: bool = False


@dataclass
class ComparisonStats:
    additions: int = 0
    deletions: int = 0
    replacements: int = 0
    moved_passages: int = 0
    paired_sheets: int = 0


@dataclass
class ComparisonResult:
    output_pdf: Path
    stats: ComparisonStats
    backend: str
    warnings: list[str] = field(default_factory=list)


def _run(command: Sequence[str], *, cwd: Path | None = None) -> subprocess.CompletedProcess[str]:
    try:
        return subprocess.run(
            list(command),
            cwd=str(cwd) if cwd else None,
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or exc.stdout or "conversion command failed").strip()
        raise ComparisonError(detail[-1200:]) from exc


def _find_soffice() -> Path | None:
    found = shutil.which("soffice") or shutil.which("soffice.exe")
    if found:
        return Path(found)
    candidates = [
        Path(os.environ.get("PROGRAMFILES", r"C:\Program Files")) / "LibreOffice/program/soffice.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")) / "LibreOffice/program/soffice.exe",
    ]
    return next((path for path in candidates if path.exists()), None)


def _word_available() -> bool:
    if os.name != "nt":
        return False
    command = [
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        "$w=$null; try {$w=New-Object -ComObject Word.Application; exit 0} catch {exit 1} finally {if($w){$w.Quit()}}",
    ]
    try:
        subprocess.run(command, check=True, capture_output=True, timeout=20, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return True
    except (subprocess.SubprocessError, OSError):
        return False


def choose_backend(requested: str) -> str:
    requested = requested.casefold()
    if requested not in {"auto", "word", "libreoffice"}:
        raise ComparisonError(f"Unknown conversion backend: {requested}")
    if requested == "word":
        if not _word_available():
            raise ComparisonError("Microsoft Word desktop automation is unavailable.")
        return "word"
    if requested == "libreoffice":
        if not _find_soffice():
            raise ComparisonError("LibreOffice was not found. Install it or choose Microsoft Word.")
        return "libreoffice"
    if _word_available():
        return "word"
    if _find_soffice():
        return "libreoffice"
    raise ComparisonError("Neither Microsoft Word desktop nor LibreOffice was found.")


def _convert_with_word(source: Path, destination: Path, helper: Path) -> None:
    _run([
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(helper),
        str(source.resolve()),
        str(destination.resolve()),
    ])
    if not destination.exists() or destination.stat().st_size == 0:
        raise ComparisonError("Microsoft Word did not create the expected PDF.")


def _convert_with_libreoffice(source: Path, destination: Path, workdir: Path) -> None:
    soffice = _find_soffice()
    if not soffice:
        raise ComparisonError("LibreOffice was not found.")
    outdir = workdir / "lo_output"
    profile = workdir / "lo_profile"
    outdir.mkdir(parents=True, exist_ok=True)
    profile.mkdir(parents=True, exist_ok=True)
    profile_uri = profile.resolve().as_uri()
    _run([
        str(soffice),
        "--headless",
        f"-env:UserInstallation={profile_uri}",
        "--convert-to",
        "pdf:writer_pdf_Export",
        "--outdir",
        str(outdir),
        str(source.resolve()),
    ])
    created = outdir / f"{source.stem}.pdf"
    if not created.exists():
        raise ComparisonError("LibreOffice did not create the expected PDF.")
    shutil.move(str(created), str(destination))


def convert_docx(source: Path, destination: Path, backend: str, helper: Path, workdir: Path) -> None:
    if source.suffix.casefold() != ".docx" or not source.is_file():
        raise ComparisonError(f"Not a readable DOCX file: {source}")
    if backend == "word":
        _convert_with_word(source, destination, helper)
    else:
        _convert_with_libreoffice(source, destination, workdir)


def _median(values: Iterable[float], default: float) -> float:
    ordered = sorted(values)
    if not ordered:
        return default
    middle = len(ordered) // 2
    return ordered[middle] if len(ordered) % 2 else (ordered[middle - 1] + ordered[middle]) / 2


def _group_lines(words: list[Word], page_width: float) -> list[list[Word]]:
    baseline_rows: list[list[Word]] = []
    for word in sorted(words, key=lambda item: (item.top, item.x0)):
        target = None
        for line in reversed(baseline_rows[-6:]):
            line_top = sum(item.top for item in line) / len(line)
            tolerance = max(2.2, _median((item.bottom - item.top for item in line), 10.0) * 0.30)
            if abs(word.top - line_top) <= tolerance:
                target = line
                break
        if target is None:
            baseline_rows.append([word])
        else:
            target.append(word)

    # A shared baseline does not mean a shared text line. In journal layouts, the
    # left and right columns commonly have words at exactly the same y position.
    # Split each baseline row at a column-sized horizontal gap before paragraph
    # construction; otherwise two unrelated columns become one interleaved block.
    lines: list[list[Word]] = []
    for row in baseline_rows:
        ordered = sorted(row, key=lambda item: item.x0)
        positive_gaps = [
            following.x0 - previous.x1
            for previous, following in zip(ordered, ordered[1:])
            if 0 <= following.x0 - previous.x1 <= 18
        ]
        ordinary_gap = _median(positive_gaps, 3.5)
        split_gap = max(18.0, min(page_width * 0.075, ordinary_gap * 4.5))
        segment = [ordered[0]]
        for previous, following in zip(ordered, ordered[1:]):
            if following.x0 - previous.x1 > split_gap:
                lines.append(segment)
                segment = [following]
            else:
                segment.append(following)
        lines.append(segment)
    return sorted(lines, key=lambda line: (min(item.top for item in line), min(item.x0 for item in line)))


def _order_lines_for_reading(lines: list[list[Word]], page_width: float) -> list[list[Word]]:
    """Return one-column or column-major reading order for extracted line segments."""
    if not lines:
        return []
    midpoint = page_width / 2
    tolerance = max(8.0, page_width * 0.02)
    left: list[list[Word]] = []
    right: list[list[Word]] = []
    spanning: list[list[Word]] = []
    for line in lines:
        x0 = min(word.x0 for word in line)
        x1 = max(word.x1 for word in line)
        if x1 <= midpoint + tolerance:
            left.append(line)
        elif x0 >= midpoint - tolerance:
            right.append(line)
        else:
            spanning.append(line)

    # Require sustained evidence on both sides so ordinary indented prose is not
    # mistaken for a two-column page.
    if len(left) < 3 or len(right) < 3:
        return sorted(lines, key=lambda line: (min(word.top for word in line), min(word.x0 for word in line)))

    key = lambda line: (min(word.top for word in line), min(word.x0 for word in line))
    left.sort(key=key)
    right.sort(key=key)
    spanning.sort(key=key)
    ordered: list[list[Word]] = []
    left_index = right_index = 0
    for separator in spanning:
        separator_top = min(word.top for word in separator)
        while left_index < len(left) and min(word.top for word in left[left_index]) < separator_top:
            ordered.append(left[left_index])
            left_index += 1
        while right_index < len(right) and min(word.top for word in right[right_index]) < separator_top:
            ordered.append(right[right_index])
            right_index += 1
        ordered.append(separator)
    ordered.extend(left[left_index:])
    ordered.extend(right[right_index:])
    return ordered


def _lines_to_blocks(lines: list[list[Word]], page: int) -> list[Block]:
    if not lines:
        return []
    heights = [max(word.bottom for word in line) - min(word.top for word in line) for line in lines]
    typical_height = _median(heights, 10.0)
    typical_font_size = _median((word.size for line in lines for word in line if word.size > 0), 10.0)

    def heading_like(line: list[Word]) -> bool:
        sizes = [word.size for word in line if word.size > 0]
        line_size = _median(sizes, typical_font_size)
        bold_share = sum("bold" in word.fontname.casefold() for word in line) / max(1, len(line))
        return line_size >= typical_font_size + 1.25 or (bold_share >= 0.70 and len(line) <= 18)

    blocks: list[list[list[Word]]] = []
    current: list[list[Word]] = []
    for line in lines:
        if not current:
            current = [line]
            continue
        previous = current[-1]
        gap = min(word.top for word in line) - max(word.bottom for word in previous)
        left_shift = abs(min(word.x0 for word in line) - min(word.x0 for word in previous))
        new_paragraph = gap > typical_height * 0.85 or left_shift > 42
        if heading_like(previous) != heading_like(line):
            new_paragraph = True
        if new_paragraph:
            blocks.append(current)
            current = [line]
        else:
            current.append(line)
    if current:
        blocks.append(current)
    result: list[Block] = []
    for index, block_lines in enumerate(blocks):
        block_words = [word for line in block_lines for word in line]
        text = " ".join(word.text for word in block_words)
        if any(word.token for word in block_words):
            result.append(Block(page=page, index=index, words=block_words, text=text))
    return result


def extract_pdf(pdf_path: Path) -> list[PageData]:
    pages: list[PageData] = []
    global_index = 0
    with pdfplumber.open(str(pdf_path)) as document:
        for page_index, page in enumerate(document.pages):
            raw_words = page.extract_words(
                use_text_flow=False,
                keep_blank_chars=False,
                extra_attrs=["fontname", "size"],
            ) or []
            words = [
                Word(
                    text=str(item.get("text", "")),
                    x0=float(item["x0"]),
                    top=float(item["top"]),
                    x1=float(item["x1"]),
                    bottom=float(item["bottom"]),
                    page=page_index,
                    fontname=str(item.get("fontname", "")),
                    size=float(item.get("size", 0.0) or 0.0),
                )
                for item in raw_words
                if str(item.get("text", "")).strip()
            ]
            lines = _group_lines(words, float(page.width))
            blocks = _lines_to_blocks(_order_lines_for_reading(lines, float(page.width)), page_index)
            for block in blocks:
                block.global_index = global_index
                global_index += 1
            pages.append(PageData(page_index + 1, float(page.width), float(page.height), words, blocks))
    if not pages:
        raise ComparisonError(f"No pages were extracted from {pdf_path.name}.")
    return pages


def _block_similarity(left: Block, right: Block, left_total: int, right_total: int) -> float:
    a, b = left.normalized, right.normalized
    if not a or not b:
        return 0.0
    length_ratio = min(len(a), len(b)) / max(len(a), len(b))
    if length_ratio < 0.22:
        return 0.0
    sequence = SequenceMatcher(None, a, b, autojunk=False).ratio()
    aset, bset = set(left.tokens), set(right.tokens)
    jaccard = len(aset & bset) / max(1, len(aset | bset))
    left_position = left.global_index / max(1, left_total - 1)
    right_position = right.global_index / max(1, right_total - 1)
    proximity = max(0.0, 1.0 - abs(left_position - right_position))
    return 0.65 * sequence + 0.30 * jaccard + 0.05 * proximity


def match_blocks(left_pages: list[PageData], right_pages: list[PageData]) -> tuple[list[BlockMatch], list[Block], list[Block]]:
    left_blocks = [block for page in left_pages for block in page.blocks]
    right_blocks = [block for page in right_pages for block in page.blocks]
    candidates: list[tuple[float, int, int]] = []
    for left_index, left in enumerate(left_blocks):
        for right_index, right in enumerate(right_blocks):
            score = _block_similarity(left, right, len(left_blocks), len(right_blocks))
            minimum = 0.34 if min(len(left.tokens), len(right.tokens)) >= 5 else 0.52
            if score >= minimum:
                candidates.append((score, left_index, right_index))
    candidates.sort(reverse=True)
    used_left: set[int] = set()
    used_right: set[int] = set()
    matches: list[BlockMatch] = []
    for score, left_index, right_index in candidates:
        if left_index in used_left or right_index in used_right:
            continue
        left, right = left_blocks[left_index], right_blocks[right_index]
        left_pos = left.global_index / max(1, len(left_blocks) - 1)
        right_pos = right.global_index / max(1, len(right_blocks) - 1)
        moved = (
            score >= 0.78
            and min(len(left.tokens), len(right.tokens)) >= 8
            and abs(left.global_index - right.global_index) >= 3
            and abs(left_pos - right_pos) >= 0.12
        )
        matches.append(BlockMatch(left, right, score, moved))
        used_left.add(left_index)
        used_right.add(right_index)

    # Detect local reordering as well as large absolute position changes. The
    # longest increasing sequence represents blocks that retained reading order;
    # high-confidence blocks outside it are likely moved passages.
    ordered_matches = sorted(matches, key=lambda match: match.left.global_index)
    tails: list[int] = []
    tail_match_indices: list[int] = []
    predecessors = [-1] * len(ordered_matches)
    for index, match in enumerate(ordered_matches):
        value = match.right.global_index
        position = bisect_left(tails, value)
        if position == len(tails):
            tails.append(value)
            tail_match_indices.append(index)
        else:
            tails[position] = value
            tail_match_indices[position] = index
        if position > 0:
            predecessors[index] = tail_match_indices[position - 1]
    in_order: set[int] = set()
    if tail_match_indices:
        cursor = tail_match_indices[-1]
        while cursor >= 0:
            in_order.add(cursor)
            cursor = predecessors[cursor]
    for index, match in enumerate(ordered_matches):
        if (
            index not in in_order
            and match.score >= 0.78
            and min(len(match.left.tokens), len(match.right.tokens)) >= 8
        ):
            match.moved = True

    unmatched_left = [block for index, block in enumerate(left_blocks) if index not in used_left]
    unmatched_right = [block for index, block in enumerate(right_blocks) if index not in used_right]
    return matches, unmatched_left, unmatched_right


def _word_rect(word: Word) -> tuple[float, float, float, float]:
    return (word.x0 - 0.8, word.top - 0.8, word.x1 + 0.8, word.bottom + 0.8)


def build_marks(
    matches: list[BlockMatch], unmatched_left: list[Block], unmatched_right: list[Block]
) -> tuple[list[Mark], list[Mark], ComparisonStats]:
    """Build highlights from a document-wide word alignment.

    Paragraph matches remain useful for page pairing and moved-passage flags, but
    they are intentionally not the unit of textual diffing. Renderers can split
    the same paragraph at different column or page boundaries. Flattening the
    blocks back into reading order prevents those layout-only splits from
    turning unchanged words into whole red/green passages.
    """
    left_marks: list[Mark] = []
    right_marks: list[Mark] = []
    stats = ComparisonStats()

    left_blocks = sorted([match.left for match in matches] + unmatched_left, key=lambda block: block.global_index)
    right_blocks = sorted([match.right for match in matches] + unmatched_right, key=lambda block: block.global_index)
    left_words = [word for block in left_blocks for word in block.words if word.token]
    right_words = [word for block in right_blocks for word in block.words if word.token]

    def word_key(word: Word) -> tuple[int, float, float, float, float]:
        return (word.page, word.x0, word.top, word.x1, word.bottom)

    moved_left: set[tuple[int, float, float, float, float]] = set()
    moved_right: set[tuple[int, float, float, float, float]] = set()
    for match in matches:
        if match.moved:
            moved_left.update(word_key(word) for word in match.left.words)
            moved_right.update(word_key(word) for word in match.right.words)

    matcher = SequenceMatcher(
        None,
        [word.token for word in left_words],
        [word.token for word in right_words],
        autojunk=True,
    )
    left_changes: dict[tuple[int, float, float, float, float], str] = {}
    right_changes: dict[tuple[int, float, float, float, float], str] = {}
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        changed_left = [word for word in left_words[i1:i2] if word_key(word) not in moved_left]
        changed_right = [word for word in right_words[j1:j2] if word_key(word) not in moved_right]
        if tag == "delete":
            left_changes.update((word_key(word), "deleted") for word in changed_left)
        elif tag == "insert":
            right_changes.update((word_key(word), "added") for word in changed_right)
        elif tag == "replace":
            left_changes.update((word_key(word), "replaced") for word in changed_left)
            right_changes.update((word_key(word), "replaced") for word in changed_right)

    # A long document can repeat the same phrase in several places. The primary
    # alignment may anchor one occurrence and classify another identical
    # occurrence as a deletion/addition. Reconcile the remaining changed words
    # and remove shared runs of three or more words. This keeps genuine local
    # edits while eliminating duplicate-text false positives.
    marked_left = [word for word in left_words if word_key(word) in left_changes]
    marked_right = [word for word in right_words if word_key(word) in right_changes]
    reconciliation = SequenceMatcher(
        None,
        [word.token for word in marked_left],
        [word.token for word in marked_right],
        autojunk=False,
    )
    for block in reconciliation.get_matching_blocks():
        if block.size < 3:
            continue
        for word in marked_left[block.a:block.a + block.size]:
            left_changes.pop(word_key(word), None)
        for word in marked_right[block.b:block.b + block.size]:
            right_changes.pop(word_key(word), None)

    # Refine each corresponding paragraph locally. This resolves repeated-word
    # ambiguity in heavily revised paragraphs: even when the document-wide
    # matcher chose a different occurrence as its anchor, equal words inside a
    # high-confidence paragraph pair are unchanged and must not be highlighted.
    for match in matches:
        if match.score < 0.50 or match.moved:
            continue
        local_left = [word for word in match.left.words if word.token]
        local_right = [word for word in match.right.words if word.token]
        local_matcher = SequenceMatcher(
            None,
            [word.token for word in local_left],
            [word.token for word in local_right],
            autojunk=False,
        )
        for tag, i1, i2, j1, j2 in local_matcher.get_opcodes():
            if tag != "equal":
                continue
            for word in local_left[i1:i2]:
                left_changes.pop(word_key(word), None)
            for word in local_right[j1:j2]:
                right_changes.pop(word_key(word), None)

    # Finally reconcile identical changed phrases without requiring the phrases
    # to appear in the same document order. This covers paragraphs moved across
    # a column or page boundary. Only contiguous source runs of at least three
    # words qualify, which avoids cancelling isolated common words.
    def changed_runs(
        words: list[Word],
        changes: dict[tuple[int, float, float, float, float], str],
    ) -> list[list[Word]]:
        runs: list[list[Word]] = []
        current: list[Word] = []
        for word in words:
            if word_key(word) in changes:
                current.append(word)
            elif current:
                runs.append(current)
                current = []
        if current:
            runs.append(current)
        return runs

    left_runs = changed_runs(left_words, left_changes)
    right_runs = changed_runs(right_words, right_changes)
    phrase_candidates: list[tuple[int, int, int, int, int]] = []
    for left_run_index, left_run in enumerate(left_runs):
        left_tokens = [word.token for word in left_run]
        if len(left_tokens) < 3:
            continue
        for right_run_index, right_run in enumerate(right_runs):
            right_tokens = [word.token for word in right_run]
            if len(right_tokens) < 3:
                continue
            right_triples: dict[tuple[str, str, str], list[int]] = {}
            for right_start in range(len(right_tokens) - 2):
                right_triples.setdefault(tuple(right_tokens[right_start:right_start + 3]), []).append(right_start)
            for left_start in range(len(left_tokens) - 2):
                triple = tuple(left_tokens[left_start:left_start + 3])
                for right_start in right_triples.get(triple, []):
                    length = 3
                    while (
                        left_start + length < len(left_tokens)
                        and right_start + length < len(right_tokens)
                        and left_tokens[left_start + length] == right_tokens[right_start + length]
                    ):
                        length += 1
                    phrase_candidates.append((length, left_run_index, left_start, right_run_index, right_start))
    phrase_candidates.sort(reverse=True)
    used_left_words: set[tuple[int, int]] = set()
    used_right_words: set[tuple[int, int]] = set()
    for length, left_run_index, left_start, right_run_index, right_start in phrase_candidates:
        left_slots = {(left_run_index, index) for index in range(left_start, left_start + length)}
        right_slots = {(right_run_index, index) for index in range(right_start, right_start + length)}
        if left_slots & used_left_words or right_slots & used_right_words:
            continue
        used_left_words.update(left_slots)
        used_right_words.update(right_slots)
        for word in left_runs[left_run_index][left_start:left_start + length]:
            left_changes.pop(word_key(word), None)
        for word in right_runs[right_run_index][right_start:right_start + length]:
            right_changes.pop(word_key(word), None)

    # Short numeric table cells need a spatial tie-breaker.  Values such as
    # ``0.5`` often occur in several rows, so sequence alignment alone can pair
    # an occurrence with an identical value in a neighbouring row.  When the
    # two rendered pages retain the same cell position, compare the numeric
    # blocks at that position and make their marks authoritative.
    def numeric_items(block: Block) -> list[tuple[Word, str]]:
        items: list[tuple[Word, str]] = []
        pending_negative = False
        for word in block.words:
            raw = word.text.strip()
            if raw in {"-", "−"}:
                pending_negative = True
                continue
            if raw == "+":
                pending_negative = False
                continue
            token = word.token
            if not token:
                continue
            if re.fullmatch(r"\d+(?:[.,]\d+)*", token) is None:
                return []
            items.append((word, ("-" if pending_negative else "") + token))
            pending_negative = False
        return items if 1 <= len(items) <= 4 else []

    def block_center(block: Block) -> tuple[float, float]:
        return (
            (min(word.x0 for word in block.words) + max(word.x1 for word in block.words)) / 2,
            (min(word.top for word in block.words) + max(word.bottom for word in block.words)) / 2,
        )

    def row_context(block: Block, all_blocks: list[Block]) -> set[str]:
        target_x = min(word.x0 for word in block.words)
        target_top = min(word.top for word in block.words)
        target_bottom = max(word.bottom for word in block.words)
        target_height = max(1.0, target_bottom - target_top)
        context: set[str] = set()
        for other in all_blocks:
            if other is block or other.page != block.page:
                continue
            other_x1 = max(word.x1 for word in other.words)
            other_top = min(word.top for word in other.words)
            other_bottom = max(word.bottom for word in other.words)
            other_height = other_bottom - other_top
            overlaps_row = min(target_bottom, other_bottom) >= max(target_top, other_top)
            if other_x1 >= target_x + 5.0 or not overlaps_row or other_height > max(30.0, target_height * 3.0):
                continue
            context.update(
                token
                for token in other.tokens
                if len(token) > 1 and re.fullmatch(r"\d+(?:[.,]\d+)*", token) is None
            )
        return context

    numeric_left = [(block, numeric_items(block), row_context(block, left_blocks)) for block in left_blocks]
    numeric_right = [(block, numeric_items(block), row_context(block, right_blocks)) for block in right_blocks]
    spatial_candidates: list[tuple[float, float, int, int]] = []
    for left_index, (left_block, left_items, left_context) in enumerate(numeric_left):
        if not left_items or not left_context:
            continue
        left_x, left_y = block_center(left_block)
        for right_index, (right_block, right_items, right_context) in enumerate(numeric_right):
            shared_context = left_context & right_context
            if (
                not right_items
                or not right_context
                or not shared_context
                or left_block.page != right_block.page
            ):
                continue
            right_x, right_y = block_center(right_block)
            x_distance = abs(left_x - right_x)
            y_distance = abs(left_y - right_y)
            # Wrapped row labels can shift the same table row by one or two
            # text lines between drafts.  The matching row context prevents a
            # nearby but different row from qualifying.
            if x_distance <= 35.0 and y_distance <= 25.0:
                context_similarity = len(shared_context) / len(left_context | right_context)
                spatial_candidates.append(
                    (1.0 - context_similarity, y_distance * 10.0 + x_distance, left_index, right_index)
                )
    spatial_candidates.sort()
    used_spatial_left: set[int] = set()
    used_spatial_right: set[int] = set()
    for _, _, left_index, right_index in spatial_candidates:
        if left_index in used_spatial_left or right_index in used_spatial_right:
            continue
        used_spatial_left.add(left_index)
        used_spatial_right.add(right_index)
        left_items = numeric_left[left_index][1]
        right_items = numeric_right[right_index][1]
        spatial_matcher = SequenceMatcher(
            None,
            [token for _, token in left_items],
            [token for _, token in right_items],
            autojunk=False,
        )
        for tag, i1, i2, j1, j2 in spatial_matcher.get_opcodes():
            if tag == "equal":
                for word, _ in left_items[i1:i2]:
                    left_changes.pop(word_key(word), None)
                for word, _ in right_items[j1:j2]:
                    right_changes.pop(word_key(word), None)
                continue
            for word, _ in left_items[i1:i2]:
                left_changes[word_key(word)] = "replaced"
            for word, _ in right_items[j1:j2]:
                right_changes[word_key(word)] = "replaced"

    stats.deletions = sum(kind == "deleted" for kind in left_changes.values())
    stats.additions = sum(kind == "added" for kind in right_changes.values())
    stats.replacements = max(
        sum(kind == "replaced" for kind in left_changes.values()),
        sum(kind == "replaced" for kind in right_changes.values()),
    )
    left_marks.extend(
        Mark(word.page, _word_rect(word), left_changes[word_key(word)])
        for word in left_words
        if word_key(word) in left_changes
    )
    right_marks.extend(
        Mark(word.page, _word_rect(word), right_changes[word_key(word)])
        for word in right_words
        if word_key(word) in right_changes
    )

    for match in matches:
        if not match.moved:
            continue
        stats.moved_passages += 1
        for block, marks in ((match.left, left_marks), (match.right, right_marks)):
            x0 = min(word.x0 for word in block.words)
            top = min(word.top for word in block.words)
            x1 = max(word.x1 for word in block.words)
            bottom = max(word.bottom for word in block.words)
            marks.append(Mark(block.page, (x0 - 2, top - 2, x1 + 2, bottom + 2), "moved", True))
    return left_marks, right_marks, stats


def pair_pages(left_pages: list[PageData], right_pages: list[PageData], matches: list[BlockMatch]) -> list[tuple[int | None, int | None]]:
    left_count, right_count = len(left_pages), len(right_pages)
    weights = [[0.0 for _ in range(right_count)] for _ in range(left_count)]
    totals_left = [max(1, len(page.words)) for page in left_pages]
    totals_right = [max(1, len(page.words)) for page in right_pages]
    for match in matches:
        amount = min(len(match.left.words), len(match.right.words)) * match.score
        weights[match.left.page][match.right.page] += amount
    similarity = [
        [weights[i][j] / math.sqrt(totals_left[i] * totals_right[j]) for j in range(right_count)]
        for i in range(left_count)
    ]
    gap = -0.18
    dp = [[-1e9] * (right_count + 1) for _ in range(left_count + 1)]
    move = [[""] * (right_count + 1) for _ in range(left_count + 1)]
    dp[0][0] = 0.0
    for i in range(1, left_count + 1):
        dp[i][0] = dp[i - 1][0] + gap
        move[i][0] = "left"
    for j in range(1, right_count + 1):
        dp[0][j] = dp[0][j - 1] + gap
        move[0][j] = "right"
    for i in range(1, left_count + 1):
        for j in range(1, right_count + 1):
            options = [
                (dp[i - 1][j - 1] + 0.15 + similarity[i - 1][j - 1], "both"),
                (dp[i - 1][j] + gap, "left"),
                (dp[i][j - 1] + gap, "right"),
            ]
            dp[i][j], move[i][j] = max(options, key=lambda item: item[0])
    pairs: list[tuple[int | None, int | None]] = []
    i, j = left_count, right_count
    while i or j:
        direction = move[i][j]
        if direction == "both":
            pairs.append((i - 1, j - 1))
            i -= 1
            j -= 1
        elif direction == "left":
            pairs.append((i - 1, None))
            i -= 1
        else:
            pairs.append((None, j - 1))
            j -= 1
    return list(reversed(pairs))


def render_pdf(pdf_path: Path, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    images: list[Path] = []
    document = pdfium.PdfDocument(str(pdf_path))
    try:
        for index in range(len(document)):
            page = document[index]
            try:
                rendered = page.render(scale=2.0).to_pil().convert("RGB")
                destination = output_dir / f"page-{index + 1}.png"
                rendered.save(destination, "PNG", optimize=True)
                images.append(destination)
            finally:
                page.close()
    finally:
        document.close()
    if not images:
        raise ComparisonError("PDF rendering produced no page images.")
    return images


MARK_COLORS = {
    "added": (78, 180, 104, 88),
    "deleted": (225, 81, 81, 88),
    "replaced": (244, 166, 54, 92),
    "moved": (139, 92, 246, 230),
}


def annotate_pages(
    image_paths: list[Path], pages: list[PageData], marks: list[Mark], output_dir: Path
) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    by_page: dict[int, list[Mark]] = {}
    for mark in marks:
        by_page.setdefault(mark.page, []).append(mark)
    results: list[Path] = []
    for index, image_path in enumerate(image_paths):
        with Image.open(image_path) as opened:
            image = opened.convert("RGBA")
        page = pages[index]
        sx, sy = image.width / page.width, image.height / page.height
        overlay = Image.new("RGBA", image.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay, "RGBA")
        for mark in by_page.get(index, []):
            x0, top, x1, bottom = mark.rect
            rect = (round(x0 * sx), round(top * sy), round(x1 * sx), round(bottom * sy))
            color = MARK_COLORS[mark.kind]
            if mark.outline_only:
                draw.rectangle(rect, outline=color, width=max(3, round(sx * 1.5)))
            else:
                draw.rectangle(rect, fill=color)
        combined = Image.alpha_composite(image, overlay).convert("RGB")
        destination = output_dir / f"page-{index + 1}.png"
        combined.save(destination, "PNG", optimize=True)
        results.append(destination)
    return results


def _draw_fitted_image(pdf: canvas.Canvas, image_path: Path, x: float, y: float, width: float, height: float) -> None:
    with Image.open(image_path) as image:
        image_width, image_height = image.size
    scale = min(width / image_width, height / image_height)
    drawn_width, drawn_height = image_width * scale, image_height * scale
    dx = x + (width - drawn_width) / 2
    dy = y + (height - drawn_height) / 2
    pdf.setStrokeColor(HexColor("#B8BEC7"))
    pdf.setLineWidth(0.6)
    pdf.rect(dx - 1, dy - 1, drawn_width + 2, drawn_height + 2, stroke=1, fill=0)
    pdf.drawImage(str(image_path), dx, dy, width=drawn_width, height=drawn_height, preserveAspectRatio=True, mask="auto")


def compose_pdf(
    output_pdf: Path,
    left_images: list[Path],
    right_images: list[Path],
    pairs: list[tuple[int | None, int | None]],
    left_label: str,
    right_label: str,
    stats: ComparisonStats,
) -> None:
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    page_width, page_height = landscape(A3)
    pdf = canvas.Canvas(str(output_pdf), pagesize=(page_width, page_height), pageCompression=1)
    pdf.setTitle("DOCX Draft Comparison")
    pdf.setAuthor("Local DOCX Comparison Tool")
    margin, gutter, header, footer = 26, 22, 62, 22
    panel_width = (page_width - 2 * margin - gutter) / 2
    panel_height = page_height - header - footer
    colors = [("Deleted", "#E15151"), ("Added", "#4EB468"), ("Replaced", "#F4A636"), ("Moved", "#8B5CF6")]
    for sheet_number, (left_page, right_page) in enumerate(pairs, start=1):
        pdf.setFillColor(HexColor("#15191F"))
        pdf.setFont("Helvetica-Bold", 15)
        pdf.drawString(margin, page_height - 24, "DOCX Draft Comparison")
        pdf.setFont("Helvetica", 8.5)
        cursor = margin
        legend_y = page_height - 43
        for label, color in colors:
            pdf.setFillColor(HexColor(color))
            pdf.rect(cursor, legend_y - 1, 10, 7, stroke=0, fill=1)
            pdf.setFillColor(HexColor("#303640"))
            pdf.drawString(cursor + 14, legend_y - 1, label)
            cursor += 64
        pdf.setFillColor(HexColor("#5D6570"))
        pdf.drawRightString(page_width - margin, page_height - 23, f"Sheet {sheet_number} of {len(pairs)}")
        labels = [
            (left_label, left_page, margin, left_images),
            (right_label, right_page, margin + panel_width + gutter, right_images),
        ]
        for label, page_index, x, images in labels:
            pdf.setFillColor(HexColor("#20252C"))
            pdf.setFont("Helvetica-Bold", 10.5)
            page_text = f"Original page {page_index + 1}" if page_index is not None else "No corresponding page"
            pdf.drawString(x, page_height - header + 7, label)
            pdf.setFont("Helvetica", 9)
            pdf.setFillColor(HexColor("#5D6570"))
            pdf.drawRightString(x + panel_width, page_height - header + 7, page_text)
            if page_index is None:
                pdf.setFillColor(HexColor("#F4F5F7"))
                pdf.roundRect(x, footer, panel_width, panel_height - 7, 6, stroke=0, fill=1)
                pdf.setFillColor(HexColor("#7A828E"))
                pdf.setFont("Helvetica-Oblique", 12)
                pdf.drawCentredString(x + panel_width / 2, footer + panel_height / 2, "Page exists only in the other version")
            else:
                _draw_fitted_image(pdf, images[page_index], x, footer, panel_width, panel_height - 7)
        pdf.setFillColor(HexColor("#6B7280"))
        pdf.setFont("Helvetica", 7.5)
        summary = (
            f"Words marked - deleted {stats.deletions}  added {stats.additions}  "
            f"replaced {stats.replacements}  moved passages {stats.moved_passages}"
        )
        pdf.drawCentredString(page_width / 2, 8, summary)
        pdf.showPage()
    pdf.save()


def compare_docx(
    left_docx: str | Path,
    right_docx: str | Path,
    output_pdf: str | Path,
    *,
    left_label: str = "Version 1",
    right_label: str = "Version 2",
    backend: str = "auto",
    progress: Progress | None = None,
) -> ComparisonResult:
    notify = progress or (lambda _message: None)
    left_source, right_source, destination = Path(left_docx), Path(right_docx), Path(output_pdf)
    if left_source.resolve() == right_source.resolve():
        raise ComparisonError("Choose two different source documents.")
    if destination.suffix.casefold() != ".pdf":
        raise ComparisonError("The output filename must end in .pdf.")
    if destination.resolve() in {left_source.resolve(), right_source.resolve()}:
        raise ComparisonError("The output may not overwrite a source document.")
    chosen_backend = choose_backend(backend)
    helper = Path(__file__).with_name("word_to_pdf.ps1")
    warnings = [
        "Word and LibreOffice can paginate the same DOCX differently; the PDF reflects the selected converter.",
        "Text inside raster images and some floating objects cannot receive word-level highlights.",
    ]
    with tempfile.TemporaryDirectory(prefix="docx_compare_") as temporary:
        workdir = Path(temporary)
        left_pdf, right_pdf = workdir / "left.pdf", workdir / "right.pdf"
        notify(f"Converting both documents locally with {chosen_backend}...")
        convert_docx(left_source, left_pdf, chosen_backend, helper, workdir / "left_convert")
        convert_docx(right_source, right_pdf, chosen_backend, helper, workdir / "right_convert")
        notify("Extracting page geometry and matching corresponding paragraphs...")
        left_pages, right_pages = extract_pdf(left_pdf), extract_pdf(right_pdf)
        matches, unmatched_left, unmatched_right = match_blocks(left_pages, right_pages)
        left_marks, right_marks, stats = build_marks(matches, unmatched_left, unmatched_right)
        pairs = pair_pages(left_pages, right_pages, matches)
        stats.paired_sheets = len(pairs)
        notify("Rendering pages and applying word-level highlights...")
        left_raw = render_pdf(left_pdf, workdir / "left_raw")
        right_raw = render_pdf(right_pdf, workdir / "right_raw")
        left_annotated = annotate_pages(left_raw, left_pages, left_marks, workdir / "left_annotated")
        right_annotated = annotate_pages(right_raw, right_pages, right_marks, workdir / "right_annotated")
        notify("Composing the landscape comparison PDF...")
        compose_pdf(destination, left_annotated, right_annotated, pairs, left_label, right_label, stats)
    if not destination.exists() or destination.stat().st_size == 0:
        raise ComparisonError("The comparison PDF was not created.")
    notify("Comparison complete.")
    return ComparisonResult(destination, stats, chosen_backend, warnings)

