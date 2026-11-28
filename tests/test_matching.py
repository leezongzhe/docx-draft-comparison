from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from comparator import (
    Block,
    PageData,
    Word,
    _group_lines,
    _lines_to_blocks,
    _order_lines_for_reading,
    _convert_with_word,
    build_marks,
    match_blocks,
    pair_pages,
)


def make_block(text: str, page: int, index: int, global_index: int) -> Block:
    words = []
    x = 10.0
    for item in text.split():
        words.append(Word(item, x, 10.0, x + len(item) * 5, 20.0, page))
        x += len(item) * 5 + 3
    return Block(page, index, words, text, global_index)


def test_word_marks_and_page_pairing() -> None:
    left_blocks = [
        make_block("The model included two hundred adults", 0, 0, 0),
        make_block("A stable paragraph remains unchanged here", 1, 0, 1),
    ]
    right_blocks = [
        make_block("The model included three hundred adults", 0, 0, 0),
        make_block("A stable paragraph remains unchanged here", 2, 0, 1),
    ]
    left_pages = [PageData(1, 100, 100, b.words, [b]) for b in left_blocks]
    right_pages = [
        PageData(1, 100, 100, right_blocks[0].words, [right_blocks[0]]),
        PageData(2, 100, 100, [], []),
        PageData(3, 100, 100, right_blocks[1].words, [right_blocks[1]]),
    ]
    matches, missing_left, missing_right = match_blocks(left_pages, right_pages)
    left_marks, right_marks, stats = build_marks(matches, missing_left, missing_right)
    assert len(matches) == 2
    assert stats.replacements >= 1
    assert left_marks and right_marks
    pairs = pair_pages(left_pages, right_pages, matches)
    assert pairs[0] == (0, 0)
    assert len(pairs) == 3


def test_two_columns_are_grouped_in_column_reading_order() -> None:
    words = [
        Word("left", 40, 20, 62, 30, 0), Word("column", 66, 20, 98, 30, 0),
        Word("right", 330, 20, 355, 30, 0), Word("column", 359, 20, 391, 30, 0),
        Word("continues", 40, 32, 82, 42, 0), Word("here", 86, 32, 106, 42, 0),
        Word("continues", 330, 32, 372, 42, 0), Word("there", 376, 32, 400, 42, 0),
        Word("third", 40, 44, 65, 54, 0), Word("line", 69, 44, 87, 54, 0),
        Word("third", 330, 44, 355, 54, 0), Word("line", 359, 44, 377, 54, 0),
    ]
    lines = _group_lines(words, 600)
    assert len(lines) == 6
    ordered = _order_lines_for_reading(lines, 600)
    assert [line[0].text for line in ordered] == ["left", "continues", "third", "right", "continues", "third"]
    blocks = _lines_to_blocks(ordered, 0)
    assert len(blocks) == 2
    assert blocks[0].normalized == "left column continues here third line"
    assert blocks[1].normalized == "right column continues there third line"


def test_unchanged_words_inside_edited_paragraph_are_not_marked() -> None:
    left = make_block(
        "The system records two metrics and operators report monthly service activity",
        0, 0, 0,
    )
    right = make_block(
        "The system records two metrics and operators report monthly verified service activity",
        0, 0, 0,
    )
    left_pages = [PageData(1, 600, 800, left.words, [left])]
    right_pages = [PageData(1, 600, 800, right.words, [right])]
    matches, missing_left, missing_right = match_blocks(left_pages, right_pages)
    left_marks, right_marks, stats = build_marks(matches, missing_left, missing_right)
    assert len(matches) == 1
    assert not left_marks
    assert stats.additions == 1
    marked_right = {
        word.text.casefold()
        for word in right.words
        if any(mark.rect == (word.x0 - 0.8, word.top - 0.8, word.x1 + 0.8, word.bottom + 0.8) for mark in right_marks)
    }
    assert marked_right == {"verified"}


def test_symbol_only_glyphs_are_not_compared_as_words() -> None:
    assert Word("████", 0, 0, 10, 10, 0).token == ""
    assert Word("•", 0, 0, 10, 10, 0).token == ""


def test_decimal_values_are_compared_in_full() -> None:
    assert Word("0.50", 0, 0, 10, 10, 0).token == "0.50"
    assert Word("81.34/100", 0, 0, 10, 10, 0).token == "81.34"
    assert Word("0.50", 0, 0, 10, 10, 0).token != Word("0.75", 0, 0, 10, 10, 0).token


def test_repeated_table_decimals_mark_both_changed_cells() -> None:
    left = make_block("alpha 0.5 rate 0.14 score 95.00 beta 0.5 rate 0.02 score 66", 0, 0, 0)
    right = make_block("alpha 1.0 rate 0.29 score 190.00 beta 1.5 rate 0.07 score 198.00", 0, 0, 0)
    left_pages = [PageData(1, 600, 800, left.words, [left])]
    right_pages = [PageData(1, 600, 800, right.words, [right])]
    matches, missing_left, missing_right = match_blocks(left_pages, right_pages)
    left_marks, right_marks, _ = build_marks(matches, missing_left, missing_right)
    marked_left = {
        word.text
        for word in left.words
        if any(mark.rect == (word.x0 - 0.8, word.top - 0.8, word.x1 + 0.8, word.bottom + 0.8) for mark in left_marks)
    }
    marked_right = {
        word.text
        for word in right.words
        if any(mark.rect == (word.x0 - 0.8, word.top - 0.8, word.x1 + 0.8, word.bottom + 0.8) for mark in right_marks)
    }
    assert {"0.5", "0.14", "95.00", "0.02", "66"} <= marked_left
    assert {"1.0", "0.29", "190.00", "1.5", "0.07", "198.00"} <= marked_right


def test_duplicate_numeric_table_values_are_paired_by_row_position() -> None:
    left_blocks = [
        Block(0, 0, [Word("Alpha", 200, 100, 230, 110, 0)], "Alpha", 0),
        Block(0, 1, [Word("Beta", 200, 120, 230, 130, 0)], "Beta", 1),
        Block(0, 2, [Word("0.75", 340, 100, 360, 110, 0)], "0.75", 2),
        Block(0, 3, [Word("0.5", 340, 120, 360, 130, 0)], "0.5", 3),
    ]
    right_blocks = [
        Block(0, 0, [Word("Alpha", 200, 85, 230, 95, 0)], "Alpha", 0),
        Block(0, 1, [Word("Beta", 200, 105, 230, 115, 0)], "Beta", 1),
        Block(0, 2, [Word("0.5", 340, 85, 360, 95, 0)], "0.5", 2),
        Block(0, 3, [Word("1.0", 340, 105, 360, 115, 0)], "1.0", 3),
    ]
    left_marks, right_marks, _ = build_marks([], left_blocks, right_blocks)
    assert len(left_marks) == 2
    assert len(right_marks) == 2
    assert all(mark.kind == "replaced" for mark in left_marks + right_marks)


def test_shifted_table_rows_do_not_override_equal_numeric_values() -> None:
    left_blocks = [
        Block(0, 0, [Word("Alpha", 200, 100, 230, 110, 0)], "Alpha", 0),
        Block(0, 1, [Word("10", 340, 100, 350, 110, 0)], "10", 1),
        Block(0, 2, [Word("Beta", 200, 120, 230, 130, 0)], "Beta", 2),
        Block(0, 3, [Word("20", 340, 120, 350, 130, 0)], "20", 3),
    ]
    right_blocks = [
        Block(0, 0, [Word("Alpha", 200, 80, 230, 90, 0)], "Alpha", 0),
        Block(0, 1, [Word("10", 340, 80, 350, 90, 0)], "10", 1),
        Block(0, 2, [Word("Beta", 200, 100, 230, 110, 0)], "Beta", 2),
        Block(0, 3, [Word("20", 340, 100, 350, 110, 0)], "20", 3),
    ]
    left_marks, right_marks, _ = build_marks([], left_blocks, right_blocks)
    assert not left_marks
    assert not right_marks


def test_full_row_context_beats_a_nearer_shared_word() -> None:
    left_blocks = [
        make_block("students Postgraduate", 0, 0, 0),
        make_block("students Lecturer", 0, 1, 1),
        Block(0, 2, [Word("5", 340, 120, 350, 130, 0)], "5", 2),
        Block(0, 3, [Word("1", 340, 140, 350, 150, 0)], "1", 3),
    ]
    left_blocks[0].words = [Word("students", 200, 120, 235, 130, 0), Word("Postgraduate", 240, 120, 300, 130, 0)]
    left_blocks[1].words = [Word("students", 200, 140, 235, 150, 0), Word("Lecturer", 240, 140, 280, 150, 0)]
    right_blocks = [
        Block(0, 0, [Word("students", 200, 100, 235, 110, 0), Word("Postgraduate", 240, 100, 300, 110, 0)], "students Postgraduate", 0),
        Block(0, 1, [Word("students", 200, 120, 235, 130, 0), Word("Lecturer", 240, 120, 280, 130, 0)], "students Lecturer", 1),
        Block(0, 2, [Word("5", 340, 100, 350, 110, 0)], "5", 2),
        Block(0, 3, [Word("1", 340, 120, 350, 130, 0)], "1", 3),
    ]
    left_marks, right_marks, _ = build_marks([], left_blocks, right_blocks)
    assert not left_marks
    assert not right_marks


def test_heading_is_not_attached_to_following_body_paragraph() -> None:
    lines = [
        [Word("Methods", 40, 20, 90, 34, 0, "Calibri-Bold", 14.0)],
        [Word("Records", 40, 38, 100, 49, 0, "ArialMT", 10.5), Word("were", 104, 38, 126, 49, 0, "ArialMT", 10.5)],
        [Word("assigned", 40, 51, 78, 62, 0, "ArialMT", 10.5)],
    ]
    blocks = _lines_to_blocks(lines, 0)
    assert [block.normalized for block in blocks] == ["methods", "records were assigned"]


def test_sentence_wrap_does_not_split_a_paragraph() -> None:
    lines = [
        [Word("First", 40, 20, 62, 31, 0, "ArialMT", 10.5), Word("sentence.", 66, 20, 108, 31, 0, "ArialMT", 10.5)],
        [Word("Second", 40, 35, 72, 46, 0, "ArialMT", 10.5), Word("sentence", 76, 35, 114, 46, 0, "ArialMT", 10.5)],
    ]
    blocks = _lines_to_blocks(lines, 0)
    assert [block.normalized for block in blocks] == ["first sentence second sentence"]


def test_block_and_page_boundaries_do_not_create_false_changes() -> None:
    left = make_block("the shared paragraph crosses a page boundary without changing", 0, 0, 0)
    right_a = make_block("the shared paragraph crosses", 0, 0, 0)
    right_b = make_block("a page boundary without changing", 1, 0, 1)
    left_pages = [PageData(1, 600, 800, left.words, [left])]
    right_pages = [
        PageData(1, 600, 800, right_a.words, [right_a]),
        PageData(2, 600, 800, right_b.words, [right_b]),
    ]
    matches, missing_left, missing_right = match_blocks(left_pages, right_pages)
    left_marks, right_marks, stats = build_marks(matches, missing_left, missing_right)
    assert not left_marks
    assert not right_marks
    assert stats.additions == stats.deletions == stats.replacements == 0


def test_document_alignment_marks_only_the_changed_word_across_split_blocks() -> None:
    left_a = make_block("system events were collected using two methods", 0, 0, 0)
    left_b = make_block("the total duration was then calculated", 0, 1, 1)
    right = make_block(
        "system events were collected using two methods the average duration was then calculated",
        0, 0, 0,
    )
    left_pages = [PageData(1, 600, 800, left_a.words + left_b.words, [left_a, left_b])]
    right_pages = [PageData(1, 600, 800, right.words, [right])]
    matches, missing_left, missing_right = match_blocks(left_pages, right_pages)
    left_marks, right_marks, stats = build_marks(matches, missing_left, missing_right)
    assert stats.replacements == 1
    assert len(left_marks) == len(right_marks) == 1
    assert left_marks[0].kind == right_marks[0].kind == "replaced"


def test_adjacent_reordered_paragraph_is_flagged_as_moved() -> None:
    left_a = make_block("records were assigned to primary and secondary groups in equal numbers", 0, 0, 0)
    left_b = make_block("the first report contained system latency with additional metrics recorded", 0, 1, 1)
    right_b = make_block("the first report contained system latency with additional metrics recorded", 0, 0, 0)
    right_a = make_block("records were assigned to primary and secondary groups in equal numbers", 0, 1, 1)
    left_pages = [PageData(1, 600, 800, left_a.words + left_b.words, [left_a, left_b])]
    right_pages = [PageData(1, 600, 800, right_b.words + right_a.words, [right_b, right_a])]
    matches, missing_left, missing_right = match_blocks(left_pages, right_pages)
    left_marks, right_marks, stats = build_marks(matches, missing_left, missing_right)
    assert stats.moved_passages == 1
    assert any(mark.kind == "moved" for mark in left_marks)
    assert any(mark.kind == "moved" for mark in right_marks)


def test_unordered_identical_phrases_are_not_red_and_green() -> None:
    left_a = make_block("alpha introduction text", 0, 0, 0)
    left_b = make_block("this identical discussion sentence moved across the page boundary", 0, 1, 1)
    right_b = make_block("this identical discussion sentence moved across the page boundary", 0, 0, 0)
    right_a = make_block("alpha introduction text", 0, 1, 1)
    left_marks, right_marks, stats = build_marks([], [left_a, left_b], [right_b, right_a])
    assert not left_marks
    assert not right_marks
    assert stats.additions == stats.deletions == stats.replacements == 0


def test_word_conversion_retries_with_another_powershell_architecture() -> None:
    with TemporaryDirectory() as folder:
        root = Path(folder)
        source = root / "source.docx"
        destination = root / "output.pdf"
        helper = root / "word_to_pdf.ps1"
        source.write_bytes(b"synthetic")
        helper.write_text("synthetic", encoding="utf-8")
        attempts: list[str] = []

        def fake_run(command, *, cwd=None):
            attempts.append(command[0])
            if len(attempts) == 1:
                from comparator import ComparisonError
                raise ComparisonError("type library not registered")
            destination.write_bytes(b"%PDF-synthetic")

        with patch("comparator._powershell_candidates", return_value=["powershell-64", "powershell-32"]):
            with patch("comparator._run", side_effect=fake_run):
                _convert_with_word(source, destination, helper)

        assert attempts == ["powershell-64", "powershell-32"]
        assert destination.read_bytes().startswith(b"%PDF")

