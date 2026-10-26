import pytest
from word_document_server.utils.find_utils import (
    normalize_newlines,
    calculate_word_char_length,
    prepare_find_prefix,
    is_text_match,
    is_whole_word_boundary,
)


def test_normalize_newlines():
    assert normalize_newlines("") == ""
    assert normalize_newlines("hello\r\nworld") == "hello\nworld"
    assert normalize_newlines("hello\rworld") == "hello\nworld"
    assert normalize_newlines("line1\r\nline2\rline3\nline4") == "line1\nline2\nline3\nline4"


def test_calculate_word_char_length():
    assert calculate_word_char_length("") == 0
    assert calculate_word_char_length("abc") == 3
    # CRLF is 2 chars in python, but 1 char (\r) in Word
    assert calculate_word_char_length("abc\r\ndef") == 7
    assert calculate_word_char_length("a" * 300) == 300


def test_prepare_find_prefix_short():
    short_text = "This is a short string"
    prefix = prepare_find_prefix(short_text, chunk_size=200)
    assert prefix == short_text
    assert len(prefix) <= 250


def test_prepare_find_prefix_long():
    long_text = "A" * 500
    prefix = prepare_find_prefix(long_text, chunk_size=200)
    assert len(prefix) == 200
    assert prefix == "A" * 200


def test_prepare_find_prefix_strips_trailing_caret():
    # If the chunk boundary cuts right after a '^', it shouldn't leave a dangling escape
    text = ("A" * 199) + "^" + ("B" * 100)
    prefix = prepare_find_prefix(text, chunk_size=200)
    assert len(prefix) == 199
    assert not prefix.endswith("^")


def test_prepare_find_prefix_converts_newlines():
    text = "Hello\r\nWorld\nAnother\rLine" + ("X" * 300)
    prefix = prepare_find_prefix(text, chunk_size=50)
    assert "^p" in prefix
    assert "\r" not in prefix
    assert "\n" not in prefix
    assert len(prefix) <= 250


def test_prepare_find_prefix_crlf_boundary():
    # CRLF split across chunk boundary (199 chars + \r + \n)
    text = ("A" * 199) + "\r\n" + ("B" * 100)
    prefix = prepare_find_prefix(text, chunk_size=200)
    # The \r at index 199 is stripped to avoid splitting CRLF
    assert not prefix.endswith("\r")
    assert not prefix.endswith("^p")


def test_is_text_match():
    # Case insensitive
    assert is_text_match("Hello World", "hello world", match_case=False) is True
    assert is_text_match("Hello World", "hello world", match_case=True) is False
    assert is_text_match("Hello World", "Hello World", match_case=True) is True

    # Newline differences
    assert is_text_match("Line 1\r\nLine 2", "Line 1\nLine 2", match_case=True) is True
    assert is_text_match("Line 1\rLine 2", "Line 1\r\nLine 2", match_case=True) is True


def test_is_whole_word_boundary():
    # Clean boundaries (spaces, punctuation, start/end of doc)
    assert is_whole_word_boundary(None, None) is True
    assert is_whole_word_boundary(" ", " ") is True
    assert is_whole_word_boundary(".", "!") is True
    assert is_whole_word_boundary("\t", "\n") is True

    # Word characters adjacent
    assert is_whole_word_boundary("a", " ") is False
    assert is_whole_word_boundary(" ", "9") is False
    assert is_whole_word_boundary("_", " ") is False
    assert is_whole_word_boundary(" ", "_") is False
