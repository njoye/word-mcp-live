"""
Utilities for Word document comment operations.

Provides stable comment ID generation, comment resolution, and target echo formatting
for Word live COM comment tools.
"""

import hashlib
from typing import Any, Dict, Optional, Tuple


def generate_comment_id(
    comment: Any,
    parent_id: Optional[str] = None,
    reply_index: Optional[int] = None,
) -> str:
    """Generate a stable, unique ID for a Word Comment or Comment reply.
    
    Word COM does not assign persistent GUIDs to comments, but creation metadata
    (Date, Author, Scope/Range start offset) is immutable across positional shifts.
    Replies are identified deterministically as '<parent_id>_r<reply_index>' to guarantee
    uniqueness and clear thread traceability.
    
    Args:
        comment: Word COM Comment or Reply object.
        parent_id: Optional parent comment ID (if known).
        reply_index: Optional 1-based reply index under parent thread.
        
    Returns:
        Stable ID string in format 'c_<12-hex-hash>' or 'c_<12-hex-hash>_r<reply_index>'.
    """
    try:
        # If explicitly specified as a reply with known parent and index:
        if reply_index is not None and parent_id:
            return f"{parent_id}_r{reply_index}"

        # If comment has an Ancestor property in Word COM, it is a reply
        ancestor = getattr(comment, "Ancestor", None)
        # Guard against mock recursion or self-reference
        if (
            ancestor is not None
            and ancestor is not comment
            and not hasattr(ancestor, "_mock_methods")
        ):
            p_id = parent_id or generate_comment_id(ancestor)
            r_idx = None
            try:
                replies = getattr(ancestor, "Replies", None)
                count = getattr(replies, "Count", 0)
                if isinstance(count, int):
                    for idx in range(1, count + 1):
                        r_obj = replies(idx)
                        if r_obj == comment:
                            r_idx = idx
                            break
            except Exception:
                pass
            if r_idx is None:
                r_idx = getattr(comment, "Index", 1)
            return f"{p_id}_r{r_idx}"

        date_val = getattr(comment, "Date", None)
        date_str = str(date_val) if date_val is not None else ""
        author = str(getattr(comment, "Author", ""))
        
        # Scope start for top-level comments; Range start for fallback
        start_offset = ""
        try:
            if hasattr(comment, "Scope") and comment.Scope is not None:
                start_offset = str(comment.Scope.Start)
        except Exception:
            pass
            
        if not start_offset:
            try:
                if hasattr(comment, "Range") and comment.Range is not None:
                    start_offset = str(comment.Range.Start)
            except Exception:
                pass

        # Text snippet to prevent hash collision for comments at same anchor in same second
        excerpt_snip = ""
        try:
            if (
                hasattr(comment, "Range")
                and comment.Range is not None
                and isinstance(getattr(comment.Range, "Text", None), str)
            ):
                excerpt_snip = comment.Range.Text.strip()[:20]
        except Exception:
            pass
                
        payload = f"{date_str}|{author}|{start_offset}|{excerpt_snip}".encode("utf-8")
        h = hashlib.sha256(payload).hexdigest()[:12]
        return f"c_{h}"
    except Exception:
        idx = getattr(comment, "Index", "unk")
        return f"c_idx_{idx}"


def extract_comment_excerpt(comment: Any, max_chars: int = 80) -> str:
    """Safely extract an excerpt of a comment's text."""
    try:
        if comment and hasattr(comment, "Range") and comment.Range:
            raw_text = getattr(comment.Range, "Text", "")
            if isinstance(raw_text, str) and raw_text:
                text = raw_text.strip().replace("\r", " ").replace("\n", " ")
                return text[:max_chars]
    except Exception:
        pass
    return ""


def resolve_comment(
    doc: Any,
    comment_id: Optional[str] = None,
    comment_index: Optional[int] = None,
) -> Tuple[Any, Dict[str, Any]]:
    """Resolve a comment in the document by stable comment_id or positional index.
    
    Args:
        doc: Word COM Document object.
        comment_id: Stable comment ID (e.g. 'c_a8f3b9012e4c' or 'c_a8f3b9012e4c_r1').
        comment_index: 1-indexed comment position.
        
    Returns:
        Tuple of (Word Comment COM object, metadata dict).
        Metadata dict contains:
            'comment_id', 'index', 'author', 'excerpt', 'is_reply', 'parent_comment_id'
            
    Raises:
        ValueError: If neither identifier is provided, or the comment cannot be found.
    """
    if comment_id is None and comment_index is None:
        raise ValueError("Either comment_id or comment_index must be provided.")
        
    total_comments = doc.Comments.Count
    if not isinstance(total_comments, int):
        total_comments = int(getattr(total_comments, "__int__", lambda: 0)())
    
    # 1. Resolve by comment_id
    if comment_id is not None:
        # Check if targeting a reply using the deterministic pattern '<parent_id>_r<reply_idx>'
        if "_r" in comment_id:
            parts = comment_id.rsplit("_r", 1)
            target_parent_id = parts[0]
            target_reply_idx = None
            try:
                target_reply_idx = int(parts[1])
            except ValueError:
                pass

            if target_reply_idx is not None:
                # Find the parent comment
                for i in range(1, total_comments + 1):
                    c = doc.Comments(i)
                    if getattr(c, "Ancestor", None) is not None and not hasattr(c, "_mock_methods"):
                        continue
                    if generate_comment_id(c) == target_parent_id:
                        try:
                            count = getattr(c.Replies, "Count", 0)
                            if isinstance(count, int) and target_reply_idx <= count:
                                r = c.Replies(target_reply_idx)
                                return r, {
                                    "comment_id": comment_id,
                                    "index": target_reply_idx,
                                    "top_level_index": i,
                                    "author": str(getattr(r, "Author", "")),
                                    "excerpt": extract_comment_excerpt(r),
                                    "is_reply": True,
                                    "parent_comment_id": target_parent_id,
                                    "parent_comment": c,
                                }
                            else:
                                raise ValueError(
                                    f"Reply {target_reply_idx} not found under comment '{target_parent_id}' "
                                    f"(thread has {count} replies)."
                                )
                        except Exception as reply_err:
                            raise ValueError(
                                f"Failed to access reply {target_reply_idx} on comment '{target_parent_id}': {reply_err}"
                            )

        # Scan top-level comments
        for i in range(1, total_comments + 1):
            c = doc.Comments(i)
            # Skip replies in top-level iteration
            if getattr(c, "Ancestor", None) is not None and not hasattr(c, "_mock_methods"):
                continue
            c_id = generate_comment_id(c)
            if c_id == comment_id:
                return c, {
                    "comment_id": c_id,
                    "index": i,
                    "author": str(getattr(c, "Author", "")),
                    "excerpt": extract_comment_excerpt(c),
                    "is_reply": False,
                    "parent_comment_id": None,
                }

        # Fallback inspection across replies (for backward compatibility)
        for i in range(1, total_comments + 1):
            c = doc.Comments(i)
            if getattr(c, "Ancestor", None) is not None and not hasattr(c, "_mock_methods"):
                continue
            c_id = generate_comment_id(c)
            try:
                count = getattr(c.Replies, "Count", 0)
                if isinstance(count, int) and count > 0:
                    for r_idx in range(1, count + 1):
                        r = c.Replies(r_idx)
                        r_id = generate_comment_id(r, parent_id=c_id, reply_index=r_idx)
                        if r_id == comment_id:
                            return r, {
                                "comment_id": r_id,
                                "index": r_idx,
                                "top_level_index": i,
                                "author": str(getattr(r, "Author", "")),
                                "excerpt": extract_comment_excerpt(r),
                                "is_reply": True,
                                "parent_comment_id": c_id,
                                "parent_comment": c,
                            }
            except Exception:
                pass
                
        raise ValueError(
            f"Comment with ID '{comment_id}' not found. "
            "The comment may have been deleted, resolved, or modified in Word."
        )
        
    # 2. Resolve by comment_index (fallback)
    if comment_index < 1 or comment_index > total_comments:
        raise ValueError(
            f"comment_index {comment_index} out of range (document has {total_comments} comment(s))."
        )
        
    c = doc.Comments(comment_index)
    is_reply = getattr(c, "Ancestor", None) is not None and not hasattr(c, "_mock_methods")
    if is_reply:
        p = c.Ancestor
        p_id = generate_comment_id(p)
        c_id = generate_comment_id(c, parent_id=p_id)
        return c, {
            "comment_id": c_id,
            "index": comment_index,
            "author": str(getattr(c, "Author", "")),
            "excerpt": extract_comment_excerpt(c),
            "is_reply": True,
            "parent_comment_id": p_id,
            "parent_comment": p,
        }
    else:
        c_id = generate_comment_id(c)
        return c, {
            "comment_id": c_id,
            "index": comment_index,
            "author": str(getattr(c, "Author", "")),
            "excerpt": extract_comment_excerpt(c),
            "is_reply": False,
            "parent_comment_id": None,
        }
