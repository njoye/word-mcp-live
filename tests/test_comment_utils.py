"""Unit tests for comment_utils module."""

import pytest
from unittest.mock import MagicMock
from word_document_server.utils.comment_utils import (
    generate_comment_id,
    extract_comment_excerpt,
    resolve_comment,
)


def test_generate_comment_id_format():
    mock_comment = MagicMock()
    mock_comment.Date = "2026-10-26T12:00:00"
    mock_comment.Author = "Timo Mueller"
    mock_range = MagicMock()
    mock_range.Start = 42
    mock_comment.Scope = mock_range

    cid = generate_comment_id(mock_comment)
    assert cid.startswith("c_")
    assert len(cid) == 14  # 'c_' + 12 hex chars
    assert all(c in "0123456789abcdef" for c in cid[2:])


def test_generate_comment_id_deterministic():
    mock_comment1 = MagicMock()
    mock_comment1.Date = "2026-10-26T12:00:00"
    mock_comment1.Author = "Author A"
    mock_comment1.Scope.Start = 100

    mock_comment2 = MagicMock()
    mock_comment2.Date = "2026-10-26T12:00:00"
    mock_comment2.Author = "Author A"
    mock_comment2.Scope.Start = 100

    assert generate_comment_id(mock_comment1) == generate_comment_id(mock_comment2)


def test_generate_comment_id_differs_on_attr():
    mock_comment1 = MagicMock()
    mock_comment1.Date = "2026-10-26T12:00:00"
    mock_comment1.Author = "Author A"
    mock_comment1.Scope.Start = 100

    mock_comment2 = MagicMock()
    mock_comment2.Date = "2026-10-26T12:01:00"
    mock_comment2.Author = "Author A"
    mock_comment2.Scope.Start = 100

    assert generate_comment_id(mock_comment1) != generate_comment_id(mock_comment2)


def test_extract_comment_excerpt():
    mock_comment = MagicMock()
    mock_comment.Range.Text = "Hello World! This is a long comment text."

    excerpt = extract_comment_excerpt(mock_comment, max_chars=11)
    assert excerpt == "Hello World"

    mock_comment.Range.Text = "Short"
    assert extract_comment_excerpt(mock_comment) == "Short"

    mock_comment.Range.Text = None
    assert extract_comment_excerpt(mock_comment) == ""


def test_resolve_comment_by_id():
    mock_doc = MagicMock()
    mock_comment = MagicMock()
    mock_comment.Date = "2026-10-26T10:00:00"
    mock_comment.Author = "Reviewer"
    mock_comment.Scope.Start = 10
    mock_comment.Range.Text = "Comment text"
    mock_doc.Comments.Count = 1
    mock_doc.Comments.return_value = mock_comment

    cid = generate_comment_id(mock_comment)
    resolved, meta = resolve_comment(mock_doc, comment_id=cid)
    assert resolved == mock_comment
    assert meta["index"] == 1
    assert meta["comment_id"] == cid
    assert meta["author"] == "Reviewer"
    assert meta["is_reply"] is False


def test_resolve_comment_not_found():
    mock_doc = MagicMock()
    mock_doc.Comments.Count = 0
    with pytest.raises(ValueError, match="not found"):
        resolve_comment(mock_doc, comment_id="c_nonexistent")


def test_resolve_comment_by_index():
    mock_doc = MagicMock()
    mock_comment = MagicMock()
    mock_comment.Date = "2026-10-26T10:00:00"
    mock_comment.Author = "Author"
    mock_comment.Scope.Start = 5
    mock_comment.Range.Text = "Comment 2"
    mock_doc.Comments.Count = 3
    mock_doc.Comments.return_value = mock_comment

    resolved, meta = resolve_comment(mock_doc, comment_index=2)
    assert resolved == mock_comment
    assert meta["index"] == 2
    mock_doc.Comments.assert_called_with(2)


def test_resolve_comment_out_of_bounds():
    mock_doc = MagicMock()
    mock_doc.Comments.Count = 2
    with pytest.raises(ValueError, match="out of range"):
        resolve_comment(mock_doc, comment_index=5)

    with pytest.raises(ValueError, match="out of range"):
        resolve_comment(mock_doc, comment_index=0)


def test_reply_comment_id_format_and_uniqueness():
    mock_parent = MagicMock()
    mock_parent.Date = "2026-10-26T12:00:00"
    mock_parent.Author = "ParentAuthor"
    mock_parent.Scope.Start = 100
    mock_parent.Range.Text = "Parent comment text"

    parent_id = generate_comment_id(mock_parent)
    assert parent_id.startswith("c_")
    assert "_r" not in parent_id

    mock_reply1 = MagicMock()
    mock_reply2 = MagicMock()

    reply1_id = generate_comment_id(mock_reply1, parent_id=parent_id, reply_index=1)
    reply2_id = generate_comment_id(mock_reply2, parent_id=parent_id, reply_index=2)

    assert reply1_id == f"{parent_id}_r1"
    assert reply2_id == f"{parent_id}_r2"
    assert reply1_id != reply2_id
    assert reply1_id != parent_id


def test_resolve_reply_by_comment_id():
    mock_doc = MagicMock()
    mock_parent = MagicMock()
    mock_parent.Date = "2026-10-26T12:00:00"
    mock_parent.Author = "ParentAuthor"
    mock_parent.Scope.Start = 100
    mock_parent.Range.Text = "Parent comment"
    mock_parent.Ancestor = None

    mock_reply1 = MagicMock()
    mock_reply1.Author = "ReplyAuthor"
    mock_reply1.Range.Text = "First reply"

    mock_parent.Replies.Count = 2
    mock_parent.Replies.side_effect = lambda idx: mock_reply1 if idx == 1 else MagicMock()

    mock_doc.Comments.Count = 1
    mock_doc.Comments.return_value = mock_parent

    parent_id = generate_comment_id(mock_parent)
    target_reply_id = f"{parent_id}_r1"

    resolved, meta = resolve_comment(mock_doc, comment_id=target_reply_id)
    assert resolved == mock_reply1
    assert meta["comment_id"] == target_reply_id
    assert meta["index"] == 1
    assert meta["is_reply"] is True
    assert meta["parent_comment_id"] == parent_id
    assert meta["parent_comment"] == mock_parent

