"""Unit tests for new Phase 1-6 features:
- F1: Stable comment IDs & target echo
- F2: Docstring Conventions & Invariant Contracts
- F4: Inline shapes listing and replacement parameter handling
- F5: Find coordinate bridge and span deduplication
- F6: Per-occurrence and start_after text replacement
"""

import json
import pytest
from unittest.mock import MagicMock, patch

from word_document_server.tools.live_read_tools import (
    word_live_find_text,
    word_live_list_inline_shapes,
    word_live_get_comments,
    word_live_reply_to_comment,
    word_live_resolve_comment,
    word_live_delete_comment,
)
from word_document_server.tools.live_tools import (
    word_live_replace_text,
    word_live_replace_inline_image,
)


# =========================================================================
# F2: Test Conventions & Invariants in docstrings
# =========================================================================
def test_conventions_in_docstrings():
    from word_document_server.tools.live_read_tools import word_live_get_text
    from word_document_server.tools.live_tools import (
        word_live_insert_text,
        word_live_format_text,
        word_live_delete_text,
        word_live_insert_paragraphs,
    )

    tools_to_check = [
        word_live_get_text,
        word_live_find_text,
        word_live_replace_text,
        word_live_insert_text,
        word_live_format_text,
        word_live_delete_text,
        word_live_insert_paragraphs,
        word_live_list_inline_shapes,
        word_live_replace_inline_image,
        word_live_get_comments,
        word_live_reply_to_comment,
        word_live_resolve_comment,
        word_live_delete_comment,
    ]

    for tool in tools_to_check:
        doc = tool.__doc__
        assert doc is not None, f"Tool {tool.__name__} has no docstring"
        assert "CONVENTIONS & INVARIANTS:" in doc, f"Tool {tool.__name__} missing CONVENTIONS & INVARIANTS"


# =========================================================================
# F5: Find span deduplication and coordinate bridge
# =========================================================================
@pytest.mark.anyio
async def test_find_text_span_deduplication_and_coords():
    mock_app = MagicMock()
    mock_doc = MagicMock()
    mock_doc.Name = "test.docx"

    # Set up mock range and find object
    mock_rng = MagicMock()
    mock_rng.Text = "Target"
    mock_rng.Start = 100
    mock_rng.End = 106

    # Coordinate bridge mocks
    mock_doc.Range.return_value.Paragraphs.Count = 5
    mock_rng.Information.return_value = 2  # Page 2

    # Simulate Find.Execute finding 2 duplicate spans then stopping
    call_count = 0
    def mock_find_execute(**kwargs):
        nonlocal call_count
        call_count += 1
        if call_count <= 2:
            mock_rng.Start = 100
            mock_rng.End = 106
            mock_rng.Text = "Target"
            return True
        return False

    mock_rng.Find.Execute.side_effect = mock_find_execute
    mock_doc.Content.Duplicate = mock_rng

    with patch("word_document_server.core.word_com.get_word_app", return_value=mock_app), \
         patch("word_document_server.core.word_com.find_document", return_value=mock_doc):
        res_raw = await word_live_find_text(filename="test.docx", search_text="Target")
        res = json.loads(res_raw)

        assert res["success"] is True
        # Deduplication should collapse the duplicate hit to 1 match
        assert res["match_count"] == 1
        assert len(res["matches"]) == 1
        match = res["matches"][0]
        assert match["start"] == 100
        assert match["end"] == 106
        assert match["paragraph_index"] == 5
        assert match["page"] == 2


# =========================================================================
# F6: Per-occurrence replace & out-of-bounds error
# =========================================================================
@pytest.mark.anyio
async def test_replace_text_occurrence_out_of_bounds():
    mock_app = MagicMock()
    mock_doc = MagicMock()
    mock_doc.Name = "test.docx"

    # Simulate Find.Execute finding only 1 occurrence
    call_count = 0
    mock_rng = MagicMock()
    def mock_find_execute(**kwargs):
        nonlocal call_count
        call_count += 1
        if call_count == 1:
            mock_rng.Start = 10
            mock_rng.End = 15
            mock_rng.Text = "Apple"
            return True
        return False

    mock_rng.Find.Execute.side_effect = mock_find_execute
    mock_doc.Content.Duplicate = mock_rng

    with patch("word_document_server.core.word_com.get_word_app", return_value=mock_app), \
         patch("word_document_server.core.word_com.find_document", return_value=mock_doc):
        # Requesting occurrence 3 when only 1 match exists
        res_raw = await word_live_replace_text(
            filename="test.docx",
            find_text="Apple",
            replace_text="Orange",
            occurrence=3,
        )
        res = json.loads(res_raw)
        assert "error" in res
        assert "Requested occurrence 3 not found" in res["error"]


@pytest.mark.anyio
async def test_replace_text_start_after():
    mock_app = MagicMock()
    mock_doc = MagicMock()
    mock_doc.Name = "test.docx"
    mock_doc.Content.End = 1000

    mock_rng = MagicMock()
    mock_rng.Start = 250
    mock_rng.End = 255
    mock_rng.Text = "Fruit"

    mock_doc.Content.Duplicate = mock_rng
    mock_rng.Find.Execute.return_value = False

    with patch("word_document_server.core.word_com.get_word_app", return_value=mock_app), \
         patch("word_document_server.core.word_com.find_document", return_value=mock_doc):
        res_raw = await word_live_replace_text(
            filename="test.docx",
            find_text="Fruit",
            replace_text="Berry",
            start_after=200,
        )
        res = json.loads(res_raw)
        assert res["success"] is True
        mock_rng.SetRange.assert_called_with(200, 1000)


# =========================================================================
# F4: Inline shapes validation
# =========================================================================
@pytest.mark.anyio
async def test_replace_inline_image_validation():
    # 1. Missing image_path
    res1 = json.loads(await word_live_replace_inline_image(image_path=""))
    assert "error" in res1
    assert "image_path is required" in res1["error"]

    # 2. Nonexistent image_path
    res2 = json.loads(await word_live_replace_inline_image(image_path="nonexistent_file_12345.png"))
    assert "error" in res2
    assert "not found" in res2["error"]

    # 3. Shape index out of bounds
    import tempfile
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
        tf_name = tf.name

    try:
        mock_app = MagicMock()
        mock_doc = MagicMock()
        mock_doc.InlineShapes.Count = 1

        with patch("word_document_server.core.word_com.get_word_app", return_value=mock_app), \
             patch("word_document_server.core.word_com.find_document", return_value=mock_doc):
            res3 = json.loads(await word_live_replace_inline_image(
                shape_index=5,
                image_path=tf_name,
            ))
            assert "error" in res3
            assert "shape_index 5 is out of bounds" in res3["error"]
    finally:
        import os
        if os.path.exists(tf_name):
            os.remove(tf_name)


# =========================================================================
# D0: Main.py adapter keyword argument pass-through test
# =========================================================================
@pytest.mark.anyio
async def test_main_word_live_replace_text_adapter():
    from word_document_server.main import register_tools, mcp
    register_tools()
    tool = await mcp.get_tool("word_live_replace_text")
    main_replace_text = tool.fn
    
    with patch("word_document_server.tools.live_tools.word_live_replace_text") as mock_live_replace:
        mock_live_replace.return_value = json.dumps({"success": True})
        
        # Test 1: Plain replace without occurrence (previously crashed with TypeError)
        await main_replace_text(
            filename="test.docx",
            find_text="Old",
            replace_text="New",
            track_changes=True,
        )
        mock_live_replace.assert_called_with(
            filename="test.docx",
            find_text="Old",
            replace_text="New",
            match_case=False,
            match_whole_word=False,
            use_wildcards=False,
            replace_all=None,
            occurrence=None,
            start_after=None,
            track_changes=True,
        )

        # Test 2: Replace with occurrence
        await main_replace_text(
            filename="test.docx",
            find_text="Old",
            replace_text="New",
            occurrence=2,
            track_changes=True,
        )
        mock_live_replace.assert_called_with(
            filename="test.docx",
            find_text="Old",
            replace_text="New",
            match_case=False,
            match_whole_word=False,
            use_wildcards=False,
            replace_all=None,
            occurrence=2,
            start_after=None,
            track_changes=True,
        )


# =========================================================================
# D2: Replace inline image deletes old shape in place
# =========================================================================
@pytest.mark.anyio
async def test_replace_inline_image_deletes_old_shape():
    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
        tf_name = tf.name

    try:
        mock_app = MagicMock()
        mock_doc = MagicMock()
        mock_doc.Name = "test.docx"
        mock_doc.InlineShapes.Count = 2
        mock_doc.Content.End = 500

        mock_old_shape = MagicMock()
        mock_old_shape.Width = 100.0
        mock_old_shape.Height = 80.0
        mock_old_shape.Type = 3  # inline_picture
        mock_old_shape.Range.Start = 50
        mock_doc.InlineShapes.return_value = mock_old_shape

        mock_new_shape = MagicMock()
        mock_new_shape.Width = 100.0
        mock_new_shape.Height = 80.0
        mock_new_shape.Range.Start = 50
        mock_new_shape.Range.Information.return_value = 1
        mock_doc.InlineShapes.AddPicture.return_value = mock_new_shape
        mock_doc.Range.return_value.Paragraphs.Count = 3

        with patch("word_document_server.core.word_com.get_word_app", return_value=mock_app), \
             patch("word_document_server.core.word_com.find_document", return_value=mock_doc), \
             patch("word_document_server.core.word_com.undo_record"):
            res_raw = await word_live_replace_inline_image(
                filename="test.docx",
                shape_index=1,
                image_path=tf_name,
                keep_size=True,
                track_changes=True,
            )
            res = json.loads(res_raw)
            assert res["success"] is True
            assert res["action"] == "replaced"
            # Crucial AC check: old shape MUST be deleted
            mock_old_shape.Delete.assert_called_once()
            assert res["paragraph_index"] == 3
            assert res["width_pt"] == 100.0
            assert res["height_pt"] == 80.0
    finally:
        if os.path.exists(tf_name):
            os.remove(tf_name)


# =========================================================================
# D4: Delete threaded comment calls DeleteRecursively
# =========================================================================
@pytest.mark.anyio
async def test_delete_comment_threaded_calls_delete_recursively():
    mock_app = MagicMock()
    mock_doc = MagicMock()
    mock_doc.Name = "test.docx"
    mock_doc.Comments.Count = 1

    mock_comment = MagicMock()
    mock_comment.Replies.Count = 3
    mock_comment.Date = "2026-10-26T12:00:00"
    mock_comment.Author = "Author"
    mock_comment.Scope.Start = 10
    mock_comment.Range.Text = "Threaded parent"
    mock_comment.Ancestor = None

    mock_doc.Comments.return_value = mock_comment

    with patch("word_document_server.core.word_com.get_word_app", return_value=mock_app), \
         patch("word_document_server.core.word_com.find_document", return_value=mock_doc), \
         patch("word_document_server.core.word_com.undo_record"):
        
        # Test 1: delete_recursively=True (default) calls DeleteRecursively
        res1_raw = await word_live_delete_comment(
            filename="test.docx",
            comment_index=1,
            delete_recursively=True,
        )
        res1 = json.loads(res1_raw)
        assert res1["success"] is True
        mock_comment.DeleteRecursively.assert_called_once()

        # Test 2: delete_recursively=False returns clean error explaining replies exist
        res2_raw = await word_live_delete_comment(
            filename="test.docx",
            comment_index=1,
            delete_recursively=False,
        )
        res2 = json.loads(res2_raw)
        assert "error" in res2
        assert "replies and cannot be deleted alone" in res2["error"]


# =========================================================================
# D6: Get comments skips replies at top level
# =========================================================================
@pytest.mark.anyio
async def test_get_comments_skips_replies_at_top_level():
    mock_app = MagicMock()
    mock_doc = MagicMock()
    mock_doc.Name = "test.docx"
    mock_doc.Comments.Count = 2

    # Comment 1: Parent
    mock_c1 = MagicMock()
    mock_c1.Date = "2026-10-26T12:00:00"
    mock_c1.Author = "Parent"
    mock_c1.Scope.Text = "Scope text"
    mock_c1.Scope.Start = 10
    mock_c1.Range.Text = "Parent comment"
    del mock_c1.Ancestor  # Ensure mock_c1 has no Ancestor attribute

    mock_reply = MagicMock()
    mock_reply.Author = "ReplyAuthor"
    mock_reply.Date = "2026-10-26T12:00:01"
    mock_reply.Range.Text = "Reply text"
    mock_reply.Index = 2
    mock_c1.Replies.Count = 1
    mock_c1.Replies.return_value = mock_reply

    # Comment 2 in flat list: Reply to c1 (has Ancestor)
    mock_c2 = MagicMock()
    mock_c2.Ancestor = mock_c1

    mock_doc.Comments.side_effect = lambda idx: mock_c1 if idx == 1 else mock_c2

    with patch("word_document_server.core.word_com.get_word_app", return_value=mock_app), \
         patch("word_document_server.core.word_com.find_document", return_value=mock_doc):
        res_raw = await word_live_get_comments(filename="test.docx")
        res = json.loads(res_raw)
        assert res["success"] is True
        # Top level list must only have 1 thread, not 2
        assert res["comment_count"] == 1
        assert len(res["comments"]) == 1
        parent_out = res["comments"][0]
        assert parent_out["reply_count"] == 1
        assert len(parent_out["replies"]) == 1
        assert parent_out["replies"][0]["comment_id"] == f"{parent_out['comment_id']}_r1"


# =========================================================================
# D7: Docstrings contain index bases and field-code inflation warning
# =========================================================================
def test_get_text_docstrings_have_field_code_and_index_base():
    from word_document_server.tools.live_read_tools import word_live_get_text, word_live_get_page_text

    for tool in [word_live_get_text, word_live_get_page_text]:
        doc = tool.__doc__
        assert "Index bases:" in doc
        assert "Field-Code Offset Inflation:" in doc


# =========================================================================
# Follow-up Item 1: replace_all / track_changes smart defaults
# =========================================================================
@pytest.mark.anyio
async def test_word_live_replace_text_smart_defaults():
    from word_document_server.tools.live_tools import word_live_replace_text

    mock_app = MagicMock()
    mock_app.Documents.Count = 1
    mock_doc = MagicMock()
    mock_doc.Name = "test.docx"
    mock_doc.TrackRevisions = False
    mock_doc.Content.End = 50

    # Setup find mock to find 1 match then stop
    mock_rng = MagicMock()
    mock_rng.Start = 10
    mock_rng.End = 13
    mock_rng.Text = "Old"
    mock_find = MagicMock()
    # Execute returns True once then False
    mock_find.Execute.side_effect = [True, False]
    mock_rng.Find = mock_find
    mock_doc.Content.Duplicate = mock_rng

    with patch("word_document_server.core.word_com.get_word_app", return_value=mock_app), \
         patch("word_document_server.core.word_com.find_document", return_value=mock_doc):
        # 1. When replace_all is None and track_changes=True:
        # Defaults replace_all=False and succeeds without triggering infinite loop guard
        res_raw = await word_live_replace_text(
            filename="test.docx",
            find_text="Old",
            replace_text="New",
            track_changes=True,
        )
        res = json.loads(res_raw)
        assert res.get("success") is True
        assert res.get("replace_all") is False
        assert res.get("replacements") == 1

        # 2. When explicit replace_all=True and track_changes=True:
        # Guard blocks it to avoid Word COM infinite loop
        res_raw2 = await word_live_replace_text(
            filename="test.docx",
            find_text="Old",
            replace_text="New",
            replace_all=True,
            track_changes=True,
        )
        res2 = json.loads(res_raw2)
        assert "error" in res2
        assert "infinite loop" in res2["error"]

        # 3. When start_after is specified and replace_all is None:
        # Defaults replace_all=False
        mock_find.Execute.side_effect = [True, False]
        res_raw3 = await word_live_replace_text(
            filename="test.docx",
            find_text="Old",
            replace_text="New",
            start_after=5,
            track_changes=False,
        )
        res3 = json.loads(res_raw3)
        assert res3.get("success") is True
        assert res3.get("replace_all") is False


# =========================================================================
# Follow-up Item 2: COM STA Concurrency & Retry
# =========================================================================
@pytest.mark.anyio
async def test_serialize_word_com_and_retry():
    import asyncio
    from word_document_server.core.word_com import serialize_word_com

    call_order = []
    attempt_count = 0

    @serialize_word_com
    async def flaky_tool(val: str):
        nonlocal attempt_count
        call_order.append(f"start_{val}")
        attempt_count += 1
        if attempt_count == 1:
            raise RuntimeError("(-2147418111, 'Call was rejected by callee.', None, None)")
        await asyncio.sleep(0.01)
        call_order.append(f"end_{val}")
        return f"done_{val}"

    # Verify retry on callee rejection
    result = await flaky_tool("first")
    assert result == "done_first"
    assert attempt_count == 2

    # Verify serialization of concurrent calls
    @serialize_word_com
    async def task_a():
        call_order.append("start_a")
        await asyncio.sleep(0.05)
        call_order.append("end_a")
        return "a"

    @serialize_word_com
    async def task_b():
        call_order.append("start_b")
        await asyncio.sleep(0.01)
        call_order.append("end_b")
        return "b"

    call_order.clear()
    results = await asyncio.gather(task_a(), task_b())
    assert results == ["a", "b"]
    # Mutex ensures task_a completes before task_b starts
    assert call_order == ["start_a", "end_a", "start_b", "end_b"]


# =========================================================================
# Follow-up Item 3: Target Warning on Multi-Document
# =========================================================================
def test_get_document_target_warning():
    from word_document_server.core.word_com import get_document_target_warning

    mock_app = MagicMock()

    # 1. Single document open, no filename: no warning
    mock_app.Documents.Count = 1
    assert get_document_target_warning(mock_app, filename=None) is None

    # 2. Multiple documents open, filename provided: no warning
    mock_app.Documents.Count = 2
    assert get_document_target_warning(mock_app, filename="mcptest.docx") is None

    # 3. Multiple documents open, filename omitted: warning emitted
    mock_app.ActiveDocument.Name = "active_manuscript.docx"
    warning = get_document_target_warning(mock_app, filename=None)
    assert warning is not None
    assert "Multiple documents open in Word (2)" in warning
    assert "active_manuscript.docx" in warning
    assert "Pass 'filename' explicitly" in warning


