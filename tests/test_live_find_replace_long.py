import json
import pytest
from unittest.mock import patch, MagicMock

from word_document_server.tools.live_read_tools import word_live_find_text
from word_document_server.tools.live_tools import word_live_replace_text


class MockFind:
    def __init__(self, rng):
        self.rng = rng

    def ClearFormatting(self):
        pass

    def Execute(
        self,
        FindText="",
        MatchCase=False,
        MatchWholeWord=False,
        MatchWildcards=False,
        Forward=True,
        Wrap=0,
    ):
        # Microsoft Word's native FindText 255-char limit check
        if len(FindText) > 255:
            raise Exception("(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'String parameter too long.', 'wdmain11.chm', 25334, -2146822434), None)")

        doc_text = self.rng.doc.text
        search_start = self.rng.Start
        # Search from search_start forward
        haystack = doc_text[search_start:]
        needle = FindText.replace("^p", "\r").replace("^t", "\t")

        if MatchCase:
            idx = haystack.find(needle)
        else:
            idx = haystack.lower().find(needle.lower())

        if idx == -1:
            return False

        match_start = search_start + idx
        match_end = match_start + len(needle)

        if MatchWholeWord and not MatchWildcards:
            if match_start > 0 and (doc_text[match_start - 1].isalnum() or doc_text[match_start - 1] == "_"):
                return False
            if match_end < len(doc_text) and (doc_text[match_end].isalnum() or doc_text[match_end] == "_"):
                return False

        self.rng.Start = match_start
        self.rng.End = match_end
        return True


class MockRange:
    def __init__(self, doc, start=0, end=0):
        self.doc = doc
        self.Start = start
        self.End = end
        self.Find = MockFind(self)

    @property
    def Text(self):
        return self.doc.text[self.Start:self.End]

    @Text.setter
    def Text(self, value):
        # In Word COM, assigning rng.Text updates the text in the document
        before = self.doc.text[:self.Start]
        after = self.doc.text[self.End:]
        self.doc.text = before + value + after
        self.End = self.Start + len(value)

    def Collapse(self, direction=0):
        if direction == 0:  # wdCollapseEnd
            self.Start = self.End
        else:
            self.End = self.Start

    def SetRange(self, start, end):
        self.Start = start
        self.End = end

    @property
    def Duplicate(self):
        return MockRange(self.doc, self.Start, self.End)


class MockDocContent:
    def __init__(self, doc):
        self.doc = doc

    @property
    def Duplicate(self):
        return MockRange(self.doc, 0, len(self.doc.text))

    @property
    def End(self):
        return len(self.doc.text)


class MockDocument:
    def __init__(self, text=""):
        self.text = text
        self.Name = "test.docx"
        self.TrackRevisions = False

    @property
    def Content(self):
        return MockDocContent(self)

    def Range(self, start, end):
        return MockRange(self, start, end)


class MockWordApp:
    def __init__(self, doc):
        self.doc = doc
        self.UserName = "Test User"
        self.UndoRecord = MagicMock()
        self.UndoRecord.IsRecordingCustomRecord = False


@pytest.fixture
def mock_word():
    doc = MockDocument()
    app = MockWordApp(doc)
    with patch("word_document_server.core.word_com.get_word_app", return_value=app), \
         patch("word_document_server.core.word_com.find_document", return_value=doc):
        yield doc, app


@pytest.mark.anyio
async def test_find_short_text(mock_word):
    doc, _ = mock_word
    doc.text = "Hello world! This is a simple test document."
    res_raw = await word_live_find_text(search_text="simple test")
    res = json.loads(res_raw)
    assert res["success"] is True
    assert res["match_count"] == 1
    assert res["matches"][0]["start"] == 23
    assert res["matches"][0]["text"] == "simple test"


@pytest.mark.anyio
async def test_find_exact_255_chars(mock_word):
    doc, _ = mock_word
    text_255 = "A" * 255
    doc.text = f"prefix_{text_255}_suffix"
    res_raw = await word_live_find_text(search_text=text_255)
    res = json.loads(res_raw)
    assert res["success"] is True
    assert res["match_count"] == 1
    assert res["matches"][0]["text"] == text_255


@pytest.mark.anyio
async def test_find_long_text_300_chars(mock_word):
    doc, _ = mock_word
    long_text = "This is a very long text sequence. " + ("0123456789abcdef " * 18)
    assert len(long_text) > 300
    doc.text = f"Intro paragraph.\n{long_text}\nOutro paragraph."

    res_raw = await word_live_find_text(search_text=long_text)
    res = json.loads(res_raw)
    assert res["success"] is True
    assert res["match_count"] == 1
    assert res["matches"][0]["text"] == long_text


@pytest.mark.anyio
async def test_find_long_text_500_chars(mock_word):
    doc, _ = mock_word
    long_text = "UniquePrefix_" + ("X" * 480) + "_UniqueSuffix"
    assert len(long_text) == 506
    doc.text = f"Header.\n{long_text}\nFooter."

    res_raw = await word_live_find_text(search_text=long_text)
    res = json.loads(res_raw)
    assert res["success"] is True
    assert res["match_count"] == 1
    assert res["matches"][0]["text"] == long_text


@pytest.mark.anyio
async def test_find_long_text_false_positive_prefix(mock_word):
    doc, _ = mock_word
    # Two passages share the same 200-char prefix, but differ at the end
    prefix = "A" * 200
    passage1 = prefix + "WRONG_ENDING_12345"
    passage2 = prefix + "RIGHT_ENDING_67890"

    doc.text = f"First: {passage1}\nSecond: {passage2}"

    res_raw = await word_live_find_text(search_text=passage2)
    res = json.loads(res_raw)
    assert res["success"] is True
    assert res["match_count"] == 1
    assert res["matches"][0]["text"] == passage2
    assert "RIGHT_ENDING" in res["matches"][0]["text"]


@pytest.mark.anyio
async def test_find_long_text_multiple_matches(mock_word):
    doc, _ = mock_word
    long_text = "RepeatedLongString_" + ("Z" * 280)
    assert len(long_text) > 290
    doc.text = f"{long_text} intermediate text {long_text}"

    res_raw = await word_live_find_text(search_text=long_text)
    res = json.loads(res_raw)
    assert res["success"] is True
    assert res["match_count"] == 2
    assert res["matches"][0]["start"] == 0
    assert res["matches"][1]["start"] > len(long_text)


@pytest.mark.anyio
async def test_find_long_text_case_sensitivity(mock_word):
    doc, _ = mock_word
    long_text_lower = "lowercase_" + ("x" * 270)
    long_text_upper = "LOWERCASE_" + ("x" * 270)
    doc.text = f"Content: {long_text_lower}"

    # match_case=False should match
    res_ci = json.loads(await word_live_find_text(search_text=long_text_upper, match_case=False))
    assert res_ci["success"] is True
    assert res_ci["match_count"] == 1

    # match_case=True should not match
    res_cs = json.loads(await word_live_find_text(search_text=long_text_upper, match_case=True))
    assert res_cs["success"] is True
    assert res_cs["match_count"] == 0


@pytest.mark.anyio
async def test_find_wildcards_exceeding_255_returns_helpful_error(mock_word):
    long_pattern = "^p" * 150
    assert len(long_pattern) > 255
    res = json.loads(await word_live_find_text(search_text=long_pattern, use_wildcards=True))
    assert "error" in res
    assert "use_wildcards=True does not support" in res["error"]


@pytest.mark.anyio
async def test_replace_short_find_with_long_replacement(mock_word):
    doc, _ = mock_word
    doc.text = "Replace [PLACEHOLDER] here."
    long_replacement = "EXPANDED_" + ("W" * 400) + "_END"
    assert len(long_replacement) > 400

    res_raw = await word_live_replace_text(find_text="[PLACEHOLDER]", replace_text=long_replacement)
    res = json.loads(res_raw)
    assert res["success"] is True
    assert res["replacements"] == 1
    assert long_replacement in doc.text
    assert "[PLACEHOLDER]" not in doc.text


@pytest.mark.anyio
async def test_replace_long_find_with_short_replacement(mock_word):
    doc, _ = mock_word
    long_target = "TARGET_SECTION_" + ("T" * 300) + "_END"
    assert len(long_target) > 300
    doc.text = f"Before {long_target} After"

    res_raw = await word_live_replace_text(find_text=long_target, replace_text="SHORT")
    res = json.loads(res_raw)
    assert res["success"] is True
    assert res["replacements"] == 1
    assert doc.text == "Before SHORT After"


@pytest.mark.anyio
async def test_replace_long_find_with_long_replacement(mock_word):
    doc, _ = mock_word
    long_target = "OLD_CHUNK_" + ("O" * 320) + "_END"
    long_replace = "NEW_CHUNK_" + ("N" * 450) + "_END"
    doc.text = f"Start: {long_target} End."

    res_raw = await word_live_replace_text(find_text=long_target, replace_text=long_replace)
    res = json.loads(res_raw)
    assert res["success"] is True
    assert res["replacements"] == 1
    assert doc.text == f"Start: {long_replace} End."


@pytest.mark.anyio
async def test_replace_long_text_replace_all(mock_word):
    doc, _ = mock_word
    long_target = "BLOCK_" + ("B" * 260)
    doc.text = f"Item 1: {long_target}; Item 2: {long_target}; Item 3: {long_target}"

    res_raw = await word_live_replace_text(find_text=long_target, replace_text="[REPLACED]", replace_all=True)
    res = json.loads(res_raw)
    assert res["success"] is True
    assert res["replacements"] == 3
    assert doc.text == "Item 1: [REPLACED]; Item 2: [REPLACED]; Item 3: [REPLACED]"


@pytest.mark.anyio
async def test_replace_long_text_replace_first_only(mock_word):
    doc, _ = mock_word
    long_target = "BLOCK_" + ("B" * 260)
    doc.text = f"Item 1: {long_target}; Item 2: {long_target}"

    res_raw = await word_live_replace_text(find_text=long_target, replace_text="[FIRST]", replace_all=False)
    res = json.loads(res_raw)
    assert res["success"] is True
    assert res["replacements"] == 1
    assert doc.text == f"Item 1: [FIRST]; Item 2: {long_target}"


@pytest.mark.anyio
async def test_replace_wildcards_exceeding_255_returns_helpful_error(mock_word):
    long_find = "^p" * 150
    assert len(long_find) > 255
    res = json.loads(await word_live_replace_text(find_text=long_find, replace_text="new", use_wildcards=True))
    assert "error" in res
    assert "use_wildcards=True does not support" in res["error"]


@pytest.mark.anyio
async def test_find_long_text_whole_word_boundaries(mock_word):
    doc, _ = mock_word
    core_text = "CORE_" + ("C" * 260) + "_END"
    # Case 1: Embedded in letters -> whole_word should NOT match
    doc.text = f"prefix{core_text}suffix and {core_text} standalone."

    res_whole = json.loads(await word_live_find_text(search_text=core_text, whole_word=True))
    assert res_whole["success"] is True
    assert res_whole["match_count"] == 1
    # Only the second occurrence (standalone) matches
    assert res_whole["matches"][0]["start"] > len(f"prefix{core_text}suffix")

    res_not_whole = json.loads(await word_live_find_text(search_text=core_text, whole_word=False))
    assert res_not_whole["success"] is True
    assert res_not_whole["match_count"] == 2


@pytest.mark.anyio
async def test_replace_long_text_delete(mock_word):
    doc, _ = mock_word
    long_target = "REMOVE_ME_" + ("D" * 280) + "_END"
    doc.text = f"Start [{long_target}] End"

    res = json.loads(await word_live_replace_text(find_text=long_target, replace_text=""))
    assert res["success"] is True
    assert res["replacements"] == 1
    assert doc.text == "Start [] End"


@pytest.mark.anyio
async def test_replace_exact_255_chars(mock_word):
    doc, _ = mock_word
    text_255 = "E" * 255
    doc.text = f"Before {text_255} After"

    res = json.loads(await word_live_replace_text(find_text=text_255, replace_text="REPLACED_255"))
    assert res["success"] is True
    assert res["replacements"] == 1
    assert doc.text == "Before REPLACED_255 After"


@pytest.mark.anyio
async def test_replace_256_chars_boundary(mock_word):
    doc, _ = mock_word
    text_256 = "F" * 256
    doc.text = f"Before {text_256} After"

    res = json.loads(await word_live_replace_text(find_text=text_256, replace_text="REPLACED_256"))
    assert res["success"] is True
    assert res["replacements"] == 1
    assert doc.text == "Before REPLACED_256 After"


@pytest.mark.anyio
async def test_replace_1000_chars_with_1200_chars(mock_word):
    doc, _ = mock_word
    target_1000 = "START1000_" + ("K" * 982) + "_END1000"
    replace_1200 = "START1200_" + ("M" * 1182) + "_END1200"
    assert len(target_1000) == 1000
    assert len(replace_1200) == 1200
    doc.text = f"Intro. {target_1000} Outro."

    res = json.loads(await word_live_replace_text(find_text=target_1000, replace_text=replace_1200))
    assert res["success"] is True
    assert res["replacements"] == 1
    assert doc.text == f"Intro. {replace_1200} Outro."


@pytest.mark.anyio
async def test_replace_long_text_across_newlines(mock_word):
    doc, _ = mock_word
    part1 = "FIRST_PARAGRAPH_" + ("P" * 150)
    part2 = "SECOND_PARAGRAPH_" + ("Q" * 150)
    full_search = f"{part1}\n{part2}"
    assert len(full_search) > 300
    doc.text = f"Header\n{part1}\r{part2}\nFooter"

    res = json.loads(await word_live_replace_text(find_text=full_search, replace_text="MERGED_PARAGRAPH"))
    assert res["success"] is True
    assert res["replacements"] == 1
    assert "MERGED_PARAGRAPH" in doc.text
    assert part1 not in doc.text
    assert part2 not in doc.text

