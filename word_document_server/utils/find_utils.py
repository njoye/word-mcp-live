"""Utilities for handling text search and replacement in Word documents.

Handles chunking, text normalization, and boundary verification for searches
that exceed Word's native 255-character Find limit.
"""

from typing import Optional


def normalize_newlines(text: str) -> str:
    """Normalize \\r\\n and \\r to \\n for cross-platform and cross-COM string comparison."""
    if not text:
        return ""
    return text.replace("\r\n", "\n").replace("\r", "\n")


def calculate_word_char_length(text: str) -> int:
    """Calculate character count as represented in Word's story/range.
    
    In Word's object model, each paragraph break is a single carriage return (\\r, 1 char).
    Python strings often contain CRLF (\\r\\n, 2 chars).
    """
    if not text:
        return 0
    return len(text.replace("\r\n", "\r"))


def prepare_find_prefix(search_text: str, chunk_size: int = 200) -> str:
    """Prepare a search prefix chunk (<= 250 characters) safe for Word's Find.Execute.
    
    Word's Find.Execute has a hard limit of 255 characters for FindText.
    We extract an initial chunk, convert newlines to Word's '^p' escape, and ensure
    it does not end with an incomplete escape character ('^').
    """
    if not search_text:
        return ""
    
    chunk = search_text[:chunk_size]
    # Guard against splitting a CRLF across the chunk boundary
    if chunk.endswith("\r") and len(search_text) > chunk_size and search_text[chunk_size] == "\n":
        chunk = chunk[:-1]
    
    # Strip any dangling escape character at the end of the slice
    if chunk.endswith("^"):
        chunk = chunk[:-1]
        
    # Convert newlines to Word Find escape sequence ^p
    find_exec = chunk.replace("\r\n", "^p").replace("\n", "^p").replace("\r", "^p")
    
    # Ensure transformed escape sequences do not push length past Word limit (250 safe ceiling)
    if len(find_exec) > 250:
        find_exec = find_exec[:250]
        if find_exec.endswith("^"):
            find_exec = find_exec[:-1]
            
    return find_exec


def is_text_match(actual: str, target: str, match_case: bool = False) -> bool:
    """Compare actual document text with target search text, normalizing newlines."""
    norm_actual = normalize_newlines(actual)
    norm_target = normalize_newlines(target)
    if match_case:
        return norm_actual == norm_target
    return norm_actual.lower() == norm_target.lower()


def is_whole_word_boundary(char_before: Optional[str], char_after: Optional[str]) -> bool:
    """Verify that adjacent characters do not form word characters (alphanumeric or underscore)."""
    if char_before and (char_before.isalnum() or char_before == "_"):
        return False
    if char_after and (char_after.isalnum() or char_after == "_"):
        return False
    return True
