"""COM-based read, comment, and revision tools for Microsoft Word.

These tools operate on documents currently open in Word via COM automation.
They provide read access and comment/revision management for locked files
that python-docx cannot open.
"""

import json
import sys
import time
from difflib import SequenceMatcher

from word_document_server.defaults import DEFAULT_AUTHOR
# macOS JXA dispatch
_MAC_AVAILABLE = __import__('sys').platform == 'darwin'



# ---------------------------------------------------------------------------
# Paragraph snapshot helpers (in-memory, per-session)
# ---------------------------------------------------------------------------

_paragraph_snapshots: dict[str, dict] = {}


def _doc_key(doc) -> str:
    """Return a stable, case-insensitive key for an open COM document."""
    return doc.Name.lower()


def _read_paragraphs(doc) -> list[dict]:
    """Read all paragraphs from *doc* and return as a list of dicts."""
    paras = []
    for i in range(1, doc.Paragraphs.Count + 1):
        text = doc.Paragraphs(i).Range.Text.rstrip("\r\x07")
        paras.append({"index": i, "text": text})
    return paras


def _store_snapshot(doc, paragraphs: list[dict]) -> None:
    """Cache *paragraphs* for later diffing."""
    _paragraph_snapshots[_doc_key(doc)] = {
        "timestamp": time.time(),
        "paragraphs": paragraphs,
    }


def _get_snapshot(doc) -> dict | None:
    """Return the stored snapshot for *doc*, or None."""
    return _paragraph_snapshots.get(_doc_key(doc))


# ---------------------------------------------------------------------------
# Snapshot / diff tools
# ---------------------------------------------------------------------------


async def word_live_take_snapshot(filename: str = None) -> str:
    """[Windows only] Store a snapshot of the current document text for later diffing.

    Call this to set a baseline without returning the full text.
    Subsequent calls to word_live_get_diff will compare against this snapshot.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON confirmation with paragraph count and timestamp.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_get_text
        import json as _json
        result = _json.loads(mac_get_text(filename=filename))
        paras = [{"index": p["index"] + 1, "text": p["text"].rstrip("\r\x07")} for p in result["paragraphs"]]
        doc_name = (filename or "active").lower()
        _paragraph_snapshots[doc_name] = {"timestamp": time.time(), "paragraphs": paras}
        return _json.dumps({"success": True, "document": filename or "active", "paragraph_count": len(paras), "snapshot_timestamp": _paragraph_snapshots[doc_name]["timestamp"], "message": "Snapshot stored. Use word_live_get_diff to see changes."}, ensure_ascii=False)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)
        paras = _read_paragraphs(doc)
        _store_snapshot(doc, paras)
        snap = _get_snapshot(doc)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraph_count": len(paras),
            "snapshot_timestamp": snap["timestamp"] if snap else None,
            "message": "Snapshot stored. Use word_live_get_diff to see changes.",
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_diff(filename: str = None) -> str:
    """[Windows only] Return only the paragraphs that changed since the last snapshot.

    Compares the current document text against the most recent snapshot
    (created by word_live_take_snapshot). Returns added, modified, and
    deleted paragraphs with their indices.

    If no snapshot exists, returns an error prompting a snapshot first.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with keys: added, modified, deleted, unchanged_count,
        snapshot_age_seconds, and current_paragraph_count.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        snap = _get_snapshot(doc)
        if snap is None:
            return json.dumps({
                "error": "No snapshot exists for this document. "
                         "Call word_live_take_snapshot first.",
            })

        old_paras = snap["paragraphs"]
        new_paras = _read_paragraphs(doc)

        old_texts = [p["text"] for p in old_paras]
        new_texts = [p["text"] for p in new_paras]

        matcher = SequenceMatcher(None, old_texts, new_texts)

        added: list[dict] = []
        modified: list[dict] = []
        deleted: list[dict] = []
        unchanged = 0

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                unchanged += (j2 - j1)
            elif tag == "replace":
                for k in range(max(i2 - i1, j2 - j1)):
                    old_idx = i1 + k if (i1 + k) < i2 else None
                    new_idx = j1 + k if (j1 + k) < j2 else None
                    if old_idx is not None and new_idx is not None:
                        modified.append({
                            "paragraph_index": new_idx + 1,
                            "old_text": old_texts[old_idx],
                            "new_text": new_texts[new_idx],
                        })
                    elif new_idx is not None:
                        added.append({
                            "paragraph_index": new_idx + 1,
                            "text": new_texts[new_idx],
                        })
                    else:
                        deleted.append({
                            "old_paragraph_index": old_idx + 1,
                            "text": old_texts[old_idx],
                        })
            elif tag == "insert":
                for j in range(j1, j2):
                    added.append({
                        "paragraph_index": j + 1,
                        "text": new_texts[j],
                    })
            elif tag == "delete":
                for i in range(i1, i2):
                    deleted.append({
                        "old_paragraph_index": i + 1,
                        "text": old_texts[i],
                    })

        # Update snapshot to current state after diffing
        _store_snapshot(doc, new_paras)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "snapshot_age_seconds": round(time.time() - snap["timestamp"], 1),
            "current_paragraph_count": len(new_paras),
            "previous_paragraph_count": len(old_paras),
            "unchanged_count": unchanged,
            "added": added,
            "modified": modified,
            "deleted": deleted,
            "has_changes": bool(added or modified or deleted),
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_snapshot_status(filename: str = None) -> str:
    """[Windows only] Check whether a snapshot exists for the document and how old it is.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with has_snapshot, age_seconds, and paragraph_count.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)
        snap = _get_snapshot(doc)

        if snap is None:
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "has_snapshot": False,
            })

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "has_snapshot": True,
            "age_seconds": round(time.time() - snap["timestamp"], 1),
            "paragraph_count": len(snap["paragraphs"]),
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


# ---------------------------------------------------------------------------
# Original read tools
# ---------------------------------------------------------------------------


async def word_live_get_text(filename: str = None) -> str:
    """Get all text from an open Word document, paragraph by paragraph.

    For documents with more than 200 paragraphs, only the first 3 pages are
    returned along with total page count. Use word_live_get_page_text to read
    specific pages of large documents.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with paragraphs list.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_get_text, mac_get_info, mac_get_page_text
        info = json.loads(mac_get_info(filename=filename))
        total_pages = info.get("pages", 1)
        if total_pages > 5:
            result = json.loads(mac_get_page_text(filename=filename, page=1, end_page=3))
            result["truncated"] = True
            result["total_pages"] = total_pages
            result["message"] = (
                f"Document has {total_pages} pages. "
                f"Showing first 3 pages only. Use word_live_get_page_text(page=N, end_page=M) "
                f"to read specific pages."
            )
            return json.dumps(result, ensure_ascii=False)
        return mac_get_text(filename=filename)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        total_paras = doc.Paragraphs.Count

        # Large document safety cap: return first 3 pages instead of all
        if total_paras > 200:
            total_pages = doc.ComputeStatistics(2)  # wdStatisticPages
            result = json.loads(await word_live_get_page_text(filename, 1, 3))
            result["truncated"] = True
            result["total_paragraphs"] = total_paras
            result["total_pages"] = total_pages
            result["message"] = (
                f"Document has {total_paras} paragraphs across {total_pages} pages. "
                f"Showing first 3 pages only. Use word_live_get_page_text(page=N, end_page=M) "
                f"to read specific pages."
            )
            return json.dumps(result, ensure_ascii=False)

        paragraphs = []
        for i in range(1, total_paras + 1):
            text = doc.Paragraphs(i).Range.Text.rstrip("\r\x07")
            paragraphs.append({"index": i, "text": text})

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraph_count": len(paragraphs),
            "paragraphs": paragraphs,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_paragraph_format(
    filename: str = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    include_runs: bool = False,
) -> str:
    """[Windows only] Inspect paragraph formatting properties for diagnostics.

    Returns detailed formatting info for each paragraph in the range. Essential for
    debugging layout issues like unexpected page breaks (caused by keep_with_next chains),
    broken list formatting, wrong styles, or inconsistent fonts.

    Per-paragraph fields returned: index, text_preview (first 80 chars), char_start, char_end,
    style, font_name, font_size, bold, italic, alignment, space_before_pt, space_after_pt,
    line_spacing, line_spacing_rule, page_break_before, keep_with_next, keep_together.
    Also: list_type, list_level, list_string (if paragraph is in a list), highlight_color.

    When include_runs=True, each paragraph also includes a "runs" array with per-run
    formatting: text, bold, italic, font_name, font_size. Consecutive words with identical
    formatting are grouped into a single run. Useful for detecting which specific words
    are bold/italic (e.g., bold sub-clause numbers in otherwise normal text).

    Args:
        filename: Document name or path (None = active document).
        start_paragraph: First paragraph (1-indexed, required).
        end_paragraph: Last paragraph (1-indexed, defaults to start_paragraph).
        include_runs: Include per-run (word-level) formatting detail (default False).

    Returns:
        JSON with formatting details per paragraph.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_get_paragraph_format
        return mac_get_paragraph_format(filename=filename, start_paragraph=start_paragraph, end_paragraph=end_paragraph, include_runs=include_runs)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if start_paragraph is None:
        return json.dumps({"error": "start_paragraph is required (1-indexed)"})

    if end_paragraph is None:
        end_paragraph = start_paragraph

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        total_paras = doc.Paragraphs.Count
        if start_paragraph < 1 or end_paragraph > total_paras:
            return json.dumps({
                "error": f"Range {start_paragraph}-{end_paragraph} out of bounds (doc has {total_paras} paragraphs)"
            })

        ALIGN_NAMES = {0: "left", 1: "center", 2: "right", 3: "justify", 4: "distribute"}
        SPACING_RULE_NAMES = {
            0: "single", 1: "1.5_lines", 2: "double",
            3: "at_least", 4: "exactly", 5: "multiple",
        }

        results = []
        for i in range(start_paragraph, end_paragraph + 1):
            para = doc.Paragraphs(i)
            rng = para.Range
            fmt = para.Format
            text = rng.Text.rstrip("\r\x07")
            preview = text[:80] + ("..." if len(text) > 80 else "")

            info = {
                "index": i,
                "text_preview": preview,
                "char_start": rng.Start,
                "char_end": rng.End,
                "style": str(rng.Style) if rng.Style else "",
                "font_name": str(rng.Font.Name) if rng.Font.Name else "",
                "font_size": rng.Font.Size if rng.Font.Size else None,
                "bold": bool(rng.Font.Bold) if rng.Font.Bold != 9999999 else "mixed",
                "italic": bool(rng.Font.Italic) if rng.Font.Italic != 9999999 else "mixed",
                "alignment": ALIGN_NAMES.get(fmt.Alignment, str(fmt.Alignment)),
                "space_before_pt": fmt.SpaceBefore,
                "space_after_pt": fmt.SpaceAfter,
                "line_spacing": fmt.LineSpacing,
                "line_spacing_rule": SPACING_RULE_NAMES.get(fmt.LineSpacingRule, str(fmt.LineSpacingRule)),
                "page_break_before": bool(fmt.PageBreakBefore),
                "keep_with_next": bool(fmt.KeepWithNext),
                "keep_together": bool(fmt.KeepTogether),
            }

            # List info
            try:
                lf = rng.ListFormat
                if lf.ListType > 0:
                    info["list_type"] = {1: "bullet", 2: "simple_number", 3: "upper_roman",
                                          4: "lower_roman", 5: "upper_letter", 6: "lower_letter"
                                          }.get(lf.ListType, f"type_{lf.ListType}")
                    info["list_level"] = lf.ListLevelNumber
                    info["list_string"] = lf.ListString
            except Exception:
                pass

            # Highlight
            try:
                info["highlight_color"] = rng.HighlightColorIndex
            except Exception:
                pass

            # Per-run formatting (word-level detail)
            if include_runs:
                try:
                    runs = []
                    current_run = None
                    for w_idx in range(1, rng.Words.Count + 1):
                        word = rng.Words(w_idx)
                        w_bold = word.Font.Bold
                        w_italic = word.Font.Italic
                        w_fname = str(word.Font.Name) if word.Font.Name else ""
                        w_fsize = word.Font.Size if word.Font.Size else None
                        fmt = {
                            "bold": bool(w_bold) if w_bold != 9999999 else "mixed",
                            "italic": bool(w_italic) if w_italic != 9999999 else "mixed",
                            "font_name": w_fname if w_fname != "9999999" else "mixed",
                            "font_size": w_fsize if w_fsize != 9999999 else "mixed",
                        }
                        if current_run and current_run["_fmt"] == fmt:
                            current_run["text"] += word.Text
                        else:
                            if current_run:
                                del current_run["_fmt"]
                                runs.append(current_run)
                            current_run = {"text": word.Text, **fmt, "_fmt": fmt}
                    if current_run:
                        del current_run["_fmt"]
                        runs.append(current_run)
                    info["runs"] = runs
                except Exception:
                    info["runs_error"] = "Could not read word-level formatting"

            results.append(info)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs": results,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_info(filename: str = None) -> str:
    """Get document info from an open Word document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with document metadata (pages, words, paragraphs, sections, etc.).
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_get_info
        return mac_get_info(filename=filename)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        # wdStatistic constants
        WD_STAT_PAGES = 2
        WD_STAT_WORDS = 0
        WD_STAT_CHARACTERS = 3
        WD_STAT_LINES = 1

        info = {
            "name": doc.Name,
            "full_path": doc.FullName,
            "pages": doc.ComputeStatistics(WD_STAT_PAGES),
            "words": doc.ComputeStatistics(WD_STAT_WORDS),
            "characters": doc.ComputeStatistics(WD_STAT_CHARACTERS),
            "lines": doc.ComputeStatistics(WD_STAT_LINES),
            "paragraphs": doc.Paragraphs.Count,
            "sections": doc.Sections.Count,
            "tables": doc.Tables.Count,
            "comments": doc.Comments.Count,
            "track_revisions": doc.TrackRevisions,
            "saved": doc.Saved,
        }

        # Built-in properties (best effort)
        try:
            props = doc.BuiltInDocumentProperties
            info["author"] = str(props("Author").Value) if props("Author").Value else ""
            info["title"] = str(props("Title").Value) if props("Title").Value else ""
            info["subject"] = str(props("Subject").Value) if props("Subject").Value else ""
        except Exception:
            pass

        return json.dumps({"success": True, **info}, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_find_text(
    filename: str = None,
    search_text: str = "",
    match_case: bool = False,
    whole_word: bool = False,
    use_wildcards: bool = False,
    context_chars: int = 60,
    max_results: int = 50,
) -> str:
    """Find text in an open Word document.

    Supports search strings of arbitrary length (automatically handles Word's 255-character
    limit via chunked prefix matching). Supports Word special characters when use_wildcards=True:
    ^m (manual page break), ^t (tab), ^p (paragraph mark), ^s (non-breaking space), and Word wildcard syntax (patterns up to 255 chars).
    Note: whole_word is ignored when use_wildcards is True (Word limitation).

    Args:
        filename: Document name or path (None = active document).
        search_text: Text to search for. With use_wildcards=True, supports ^m, ^t, ^p, ^s and Word wildcards.
        match_case: Case-sensitive search.
        whole_word: Match whole words only (ignored when use_wildcards=True).
        use_wildcards: Enable Word wildcards and special characters (^m, ^t, ^p, ^s, etc.).
        context_chars: Characters of context before/after each match (default 60).
        max_results: Maximum number of matches to return.

    Returns:
        JSON with list of matches (position, context).
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_find_text
        return mac_find_text(filename=filename, search_text=search_text, match_case=match_case, whole_word=whole_word, use_wildcards=use_wildcards, context_chars=context_chars, max_results=max_results)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if not search_text:
        return json.dumps({"error": "search_text is required"})

    # Same control-byte hazard as replace_text: \x07 and other control
    # bytes corrupt Word's Find engine. Reject before issuing Find.Execute.
    from word_document_server.utils.text_safety import reject_control_chars
    try:
        reject_control_chars("search_text", search_text)
    except ValueError as e:
        return json.dumps({"error": str(e)})

    if use_wildcards and len(search_text) > 255:
        return json.dumps({
            "error": f"use_wildcards=True does not support search patterns longer than 255 characters (Word limitation: {len(search_text)} chars). Use use_wildcards=False for literal text search of arbitrary length."
        })

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        from word_document_server.utils.find_utils import (
            prepare_find_prefix,
            calculate_word_char_length,
            is_text_match,
            is_whole_word_boundary,
        )

        app = get_word_app()
        doc = find_document(app, filename)

        # COM-marshalling can fail intermittently after MCP reconnect on
        # property access (rng.Text, doc.Name). Wrap each access so a
        # single hiccup yields a partial result rather than aborting.
        def _safe_attr(obj, attr, default=None):
            try:
                return getattr(obj, attr)
            except Exception:
                return default

        matches = []
        partial_errors = []
        rng = doc.Content.Duplicate
        rng.Find.ClearFormatting()

        is_long = len(search_text) > 255
        if is_long:
            prefix = prepare_find_prefix(search_text, 200)
            word_char_len = calculate_word_char_length(search_text)
        else:
            prefix = search_text
            word_char_len = len(search_text)

        last_start = -1

        while len(matches) < max_results:
            try:
                found = rng.Find.Execute(
                    FindText=prefix,
                    MatchCase=match_case,
                    MatchWholeWord=whole_word if (not use_wildcards and not is_long) else False,
                    MatchWildcards=use_wildcards,
                    Forward=True,
                    Wrap=0,  # wdFindStop
                )
            except Exception as e:
                partial_errors.append(f"Find.Execute failed: {e}")
                break
            if not found:
                break

            found_start = _safe_attr(rng, "Start", -1)
            found_end = _safe_attr(rng, "End", -1)
            if found_start < 0 or found_end < 0:
                break

            # Guard against repeating the same match position
            if found_start <= last_start:
                try:
                    skip_pos = last_start + 1
                    rng.SetRange(skip_pos, skip_pos)
                except Exception:
                    break
                continue

            doc_end = _safe_attr(doc.Content, "End", found_start + word_char_len)

            if is_long:
                # Verify the full string at found_start
                match_end_candidate = min(doc_end, found_start + word_char_len)
                try:
                    verify_rng = doc.Range(found_start, match_end_candidate)
                    actual_text = _safe_attr(verify_rng, "Text", "")
                except Exception:
                    actual_text = ""

                if not is_text_match(actual_text, search_text, match_case=match_case):
                    # False positive prefix match: advance past found_start
                    last_start = found_start
                    try:
                        skip_pos = found_start + 1
                        rng.SetRange(skip_pos, skip_pos)
                    except Exception:
                        break
                    continue

                if whole_word:
                    char_before = None
                    char_after = None
                    if found_start > 0:
                        try:
                            char_before = doc.Range(found_start - 1, found_start).Text
                        except Exception:
                            pass
                    if match_end_candidate < doc_end:
                        try:
                            char_after = doc.Range(match_end_candidate, match_end_candidate + 1).Text
                        except Exception:
                            pass
                    if not is_whole_word_boundary(char_before, char_after):
                        last_start = found_start
                        try:
                            skip_pos = found_start + 1
                            rng.SetRange(skip_pos, skip_pos)
                        except Exception:
                            break
                        continue

                match_start = found_start
                match_end = match_end_candidate
                match_display_text = actual_text
            else:
                match_start = found_start
                match_end = found_end
                match_display_text = _safe_attr(rng, "Text", "<unreadable>")

            # Extract context
            try:
                context_rng = rng.Duplicate
                context_start = max(0, match_start - context_chars) if match_start >= 0 else 0
                context_end = min(doc_end, match_end + context_chars) if match_end >= 0 else context_chars
                context_rng.SetRange(context_start, context_end)
                context_text = _safe_attr(context_rng, "Text", "<unreadable>")
            except Exception as e:
                context_text = f"<context unavailable: {e}>"

            matches.append({
                "start": match_start,
                "end": match_end,
                "text": match_display_text,
                "context": context_text,
            })

            last_start = match_start

            # Move past current match
            try:
                adv_pos = match_end if match_end > match_start else match_start + 1
                rng.SetRange(adv_pos, adv_pos)
            except Exception as e:
                partial_errors.append(f"advance past match failed: {e}")
                break

        result = {
            "success": True,
            "document": _safe_attr(doc, "Name", "<unknown>"),
            "search_text": search_text,
            "match_count": len(matches),
            "matches": matches,
        }
        if partial_errors:
            result["partial_errors"] = partial_errors
        return json.dumps(result, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_comments(filename: str = None) -> str:
    """Get all comments from an open Word document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with list of comments (author, date, text, scope).
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_get_comments
        return mac_get_comments(filename=filename)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        comments = []
        for i in range(1, doc.Comments.Count + 1):
            c = doc.Comments(i)
            scope_text = ""
            try:
                scope_text = c.Scope.Text[:100] if c.Scope and c.Scope.Text else ""
            except Exception:
                pass

            # Collect replies (Word 2016+)
            replies = []
            try:
                for r_idx in range(1, c.Replies.Count + 1):
                    r = c.Replies(r_idx)
                    replies.append({
                        "index": r.Index,
                        "author": str(r.Author) if r.Author else "",
                        "date": str(r.Date) if r.Date else "",
                        "text": str(r.Range.Text) if r.Range and r.Range.Text else "",
                    })
            except Exception:
                pass  # Replies not supported in older Word versions

            comment_data = {
                "index": i,
                "author": str(c.Author) if c.Author else "",
                "date": str(c.Date) if c.Date else "",
                "text": str(c.Range.Text) if c.Range and c.Range.Text else "",
                "scope": scope_text,
            }
            if replies:
                comment_data["replies"] = replies
                comment_data["reply_count"] = len(replies)
            comments.append(comment_data)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "comment_count": len(comments),
            "comments": comments,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_comment(
    filename: str = None,
    start: int = None,
    end: int = None,
    paragraph_index: int = None,
    text: str = "",
    author: str = DEFAULT_AUTHOR,
) -> str:
    """Add a comment to an open Word document.

    Specify either start/end character positions or paragraph_index.
    If paragraph_index is given, the comment is attached to the entire paragraph.

    Args:
        filename: Document name or path (None = active document).
        start: Start character position.
        end: End character position.
        paragraph_index: 1-indexed paragraph to attach comment to.
        text: Comment text.
        author: Comment author name.

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_add_comment
        return mac_add_comment(filename=filename, start=start, end=end, paragraph_index=paragraph_index, text=text, author=author)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if not text:
        return json.dumps({"error": "Comment text is required"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Determine the range to attach the comment to
        if paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                return json.dumps({
                    "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                })
            rng = doc.Paragraphs(paragraph_index).Range
        elif start is not None and end is not None:
            rng = doc.Range(start, end)
        else:
            return json.dumps({
                "error": "Provide either start/end positions or paragraph_index"
            })

        with undo_record(app, "MCP: Add Comment"):
            # Save and restore author
            prev_author = app.UserName
            app.UserName = author
            try:
                comment = doc.Comments.Add(rng, text)
            finally:
                app.UserName = prev_author

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "comment_index": comment.Index,
            "author": author,
            "text": text[:100],
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_reply_to_comment(
    filename: str = None,
    comment_index: int = None,
    text: str = "",
    author: str = DEFAULT_AUTHOR,
) -> str:
    """Reply to an existing comment in an open Word document.

    Adds a threaded reply to a top-level comment. Requires Word 2016 or later.
    Use word_live_get_comments to find the comment_index.

    Args:
        filename: Document name or path (None = active document).
        comment_index: 1-indexed comment to reply to.
        text: Reply text.
        author: Reply author name.

    Returns:
        JSON with reply info.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"error": "word_live_reply_to_comment is not available on macOS — the AppleScript dictionary does not expose this feature"})

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if comment_index is None:
        return json.dumps({"error": "comment_index is required"})
    if not text:
        return json.dumps({"error": "Reply text is required"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if comment_index < 1 or comment_index > doc.Comments.Count:
            return json.dumps({
                "error": f"comment_index {comment_index} out of range (1-{doc.Comments.Count})"
            })

        comment = doc.Comments(comment_index)

        with undo_record(app, "MCP: Reply to Comment"):
            prev_author = app.UserName
            app.UserName = author
            try:
                reply = comment.Replies.Add(comment.Scope, text)
            except AttributeError:
                return json.dumps({
                    "error": "Comment replies require Word 2016 or later."
                })
            finally:
                app.UserName = prev_author

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "comment_index": comment_index,
            "reply_text": text[:100],
            "reply_index": reply.Index,
            "author": author,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_resolve_comment(
    filename: str = None,
    comment_index: int = None,
    resolve: bool = True,
) -> str:
    """Resolve or unresolve a comment in an open Word document.

    Marks a comment thread as resolved (Done) or re-opens it.
    Requires Word 2016 or later.

    Args:
        filename: Document name or path (None = active document).
        comment_index: 1-indexed comment to resolve/unresolve.
        resolve: True to resolve (mark done), False to unresolve (re-open).

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"error": "word_live_resolve_comment is not available on macOS — the AppleScript dictionary does not expose this feature"})

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if comment_index is None:
        return json.dumps({"error": "comment_index is required"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        if comment_index < 1 or comment_index > doc.Comments.Count:
            return json.dumps({
                "error": f"comment_index {comment_index} out of range (1-{doc.Comments.Count})"
            })

        comment = doc.Comments(comment_index)

        try:
            comment.Done = resolve
        except AttributeError:
            return json.dumps({
                "error": "Comment resolve/unresolve requires Word 2016 or later."
            })
        except Exception as e:
            if "not available" in str(e).lower():
                return json.dumps({
                    "error": "Comment.Done is not available — Word's Modern Comments "
                    "does not support programmatic resolve via COM. This is a known "
                    "Microsoft limitation with no COM workaround.",
                })
            raise

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "comment_index": comment_index,
            "resolved": resolve,
            "comment_text": str(comment.Range.Text)[:100] if comment.Range else "",
            "warning": "Comment.Done may not be visible in Modern Comments UI (M365 limitation)",
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_delete_comment(
    filename: str = None,
    comment_index: int = None,
) -> str:
    """Delete a comment from an open Word document.

    Args:
        filename: Document name or path (None = active document).
        comment_index: 1-indexed comment to delete.

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_delete_comment
        return mac_delete_comment(filename=filename, comment_index=comment_index)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if comment_index is None:
        return json.dumps({"error": "comment_index is required"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if comment_index < 1 or comment_index > doc.Comments.Count:
            return json.dumps({
                "error": f"comment_index {comment_index} out of range (1-{doc.Comments.Count})"
            })

        comment = doc.Comments(comment_index)
        comment_text = str(comment.Range.Text)[:100] if comment.Range else ""

        with undo_record(app, "MCP: Delete Comment"):
            comment.Delete()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "deleted_comment_index": comment_index,
            "deleted_comment_text": comment_text,
            "remaining_comments": doc.Comments.Count,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_list_revisions(filename: str = None) -> str:
    """List all tracked changes (revisions) in an open Word document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with list of revisions (type, author, date, text).
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_list_revisions
        return mac_list_revisions(filename=filename)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        # Revision type names
        REV_TYPES = {
            1: "insert",
            2: "delete",
            3: "property",
            4: "paragraph_number",
            5: "display_field",
            6: "reconcile",
            7: "conflict",
            8: "style",
            9: "replace",
            10: "section_property",
            11: "table_property",
            12: "cell_insert",
            13: "cell_delete",
            14: "cell_merge",
        }

        revisions = []
        for i in range(1, doc.Revisions.Count + 1):
            rev = doc.Revisions(i)
            rev_text = ""
            try:
                rev_text = rev.Range.Text[:200] if rev.Range and rev.Range.Text else ""
            except Exception:
                pass

            revisions.append({
                "index": i,
                "type": REV_TYPES.get(rev.Type, f"unknown({rev.Type})"),
                "type_id": rev.Type,
                "author": str(rev.Author) if rev.Author else "",
                "date": str(rev.Date) if rev.Date else "",
                "text": rev_text,
            })

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "revision_count": len(revisions),
            "revisions": revisions,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_accept_revisions(
    filename: str = None,
    author: str = None,
    revision_ids: list = None,
) -> str:
    """Accept tracked changes in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        author: Only accept revisions by this author.
        revision_ids: List of 1-indexed revision IDs to accept. If None + no author, accept all.

    Returns:
        JSON with count of accepted revisions.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_accept_revisions
        return mac_accept_revisions(filename=filename, author=author, revision_ids=revision_ids)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, "MCP: Accept Revisions"):
            if revision_ids is not None:
                # Accept specific revisions (process in reverse to preserve indices)
                accepted = 0
                for rid in sorted(revision_ids, reverse=True):
                    if 1 <= rid <= doc.Revisions.Count:
                        doc.Revisions(rid).Accept()
                        accepted += 1
                return json.dumps({
                    "success": True,
                    "document": doc.Name,
                    "accepted": accepted,
                    "mode": "specific_ids",
                })

            if author:
                # Accept revisions by author (iterate in reverse)
                accepted = 0
                for i in range(doc.Revisions.Count, 0, -1):
                    rev = doc.Revisions(i)
                    if str(rev.Author) == author:
                        rev.Accept()
                        accepted += 1
                return json.dumps({
                    "success": True,
                    "document": doc.Name,
                    "accepted": accepted,
                    "mode": f"by_author:{author}",
                })

            # Accept all
            total = doc.Revisions.Count
            doc.AcceptAllRevisions()
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "accepted": total,
                "mode": "all",
            })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_reject_revisions(
    filename: str = None,
    author: str = None,
    revision_ids: list = None,
) -> str:
    """Reject tracked changes in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        author: Only reject revisions by this author.
        revision_ids: List of 1-indexed revision IDs to reject. If None + no author, reject all.

    Returns:
        JSON with count of rejected revisions.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_reject_revisions
        return mac_reject_revisions(filename=filename, author=author, revision_ids=revision_ids)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, "MCP: Reject Revisions"):
            if revision_ids is not None:
                rejected = 0
                for rid in sorted(revision_ids, reverse=True):
                    if 1 <= rid <= doc.Revisions.Count:
                        doc.Revisions(rid).Reject()
                        rejected += 1
                return json.dumps({
                    "success": True,
                    "document": doc.Name,
                    "rejected": rejected,
                    "mode": "specific_ids",
                })

            if author:
                rejected = 0
                for i in range(doc.Revisions.Count, 0, -1):
                    rev = doc.Revisions(i)
                    if str(rev.Author) == author:
                        rev.Reject()
                        rejected += 1
                return json.dumps({
                    "success": True,
                    "document": doc.Name,
                    "rejected": rejected,
                    "mode": f"by_author:{author}",
                })

            total = doc.Revisions.Count
            doc.RejectAllRevisions()
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "rejected": total,
                "mode": "all",
            })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_page_text(
    filename: str = None,
    page: int = 1,
    end_page: int = None,
) -> str:
    """[Windows only] Get text from specific page(s) of an open Word document.

    Returns paragraphs on the requested page(s) with char_start/char_end offsets
    that can be passed directly to word_live_format_text, word_live_delete_text, etc.

    Uses Word's GoTo API to find page boundaries. For long legal documents, this
    is much more efficient than reading all paragraphs.

    Args:
        filename: Document name or path (None = active document).
        page: Page number to read (1-indexed, required).
        end_page: Last page to read (inclusive). If None, reads only `page`.

    Returns:
        JSON with paragraphs list, each containing index, text, char_start, char_end.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_get_page_text
        return mac_get_page_text(filename=filename, page=page, end_page=end_page)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if page < 1:
        return json.dumps({"error": "page must be >= 1"})

    if end_page is not None and end_page < page:
        return json.dumps({"error": "end_page must be >= page"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        # wdStatisticPages = 2
        total_pages = doc.ComputeStatistics(2)

        if page > total_pages:
            return json.dumps({
                "error": f"Page {page} out of range (document has {total_pages} pages)"
            })

        if end_page is None:
            end_page = page

        if end_page > total_pages:
            end_page = total_pages

        # wdGoToPage=1, wdGoToAbsolute=1
        # Get start of requested page
        page_start_range = doc.GoTo(What=1, Which=1, Count=page)
        range_start = page_start_range.Start

        # Get start of page after end_page (or end of doc)
        if end_page < total_pages:
            next_page_range = doc.GoTo(What=1, Which=1, Count=end_page + 1)
            range_end = next_page_range.Start
        else:
            range_end = doc.Content.End

        # Collect paragraphs within the page range
        paragraphs = []
        for i in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(i)
            p_start = para.Range.Start
            p_end = para.Range.End

            # Skip paragraphs entirely before our range
            if p_end <= range_start:
                continue
            # Stop once we pass our range
            if p_start >= range_end:
                break

            text = para.Range.Text.rstrip("\r\x07")
            paragraphs.append({
                "index": i,
                "text": text,
                "char_start": p_start,
                "char_end": p_end,
            })

        page_label = f"{page}" if page == end_page else f"{page}-{end_page}"
        return json.dumps({
            "success": True,
            "document": doc.Name,
            "pages": page_label,
            "total_pages": total_pages,
            "paragraph_count": len(paragraphs),
            "range_start": range_start,
            "range_end": range_end,
            "paragraphs": paragraphs,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_undo_history(
    filename: str = None,
) -> str:
    """[Windows only] Get the undo stack names from an open Word document.

    Uses Word's CommandBars to read the undo dropdown list. Each MCP tool call
    that was wrapped with undo_record will appear as "MCP: <tool name>".
    Degrades gracefully if the undo list is not accessible.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with undo_entries list (most recent first) and count.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"error": "word_live_get_undo_history is not available on macOS — the AppleScript dictionary does not expose this feature"})

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        entries = []
        try:
            # CommandBar control ID 128 = Undo split dropdown (Type=6)
            # Must specify Type=6 (msoControlSplitDropdown) — without it,
            # FindControl may return a plain button (Type=1) that lacks ListCount.
            undo_control = app.CommandBars.FindControl(Type=6, Id=128)
            if undo_control is not None:
                for i in range(1, undo_control.ListCount + 1):
                    entries.append(undo_control.List(i))
        except Exception:
            # Undocumented API — may not be available in all Word versions
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "undo_entries": [],
                "count": 0,
                "note": "Undo history not accessible in this Word version",
            })

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "undo_entries": entries,
            "count": len(entries),
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_diagnose_layout(
    filename: str = None,
) -> str:
    """[Windows only] Scan an open Word document for common layout problems.

    Performs a single-pass scan of all paragraphs and reports issues that cause
    unexpected layout behavior (empty pages, content pushed to wrong page, etc.).

    Checks performed:
    1. keep_with_next chains: 5+ consecutive paragraphs with keep_with_next=true
       push content to the next page as a block (severity: high).
    2. Heading styles on body text: Heading style on paragraphs >100 chars (medium).
    3. PageBreakBefore on non-headings: may cause unexpected blank space (medium).
    4. Manual page break count: ^m characters found via Find (info).
    5. Style summary: count of paragraphs per style, flags if >50% use Heading styles.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with issues array, style_summary dict, and issue_count.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_diagnose_layout
        return mac_diagnose_layout(filename=filename)

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        issues = []
        style_counts = {}
        total_paras = doc.Paragraphs.Count

        # Single-pass paragraph scan
        kwn_chain_start = None
        kwn_chain_len = 0

        for i in range(1, total_paras + 1):
            para = doc.Paragraphs(i)
            fmt = para.Format
            rng = para.Range
            text = rng.Text.rstrip("\r\x07")
            style_name = str(rng.Style) if rng.Style else "Unknown"

            # Style summary
            style_counts[style_name] = style_counts.get(style_name, 0) + 1

            kwn = bool(fmt.KeepWithNext)

            # Track keep_with_next chains
            if kwn:
                if kwn_chain_start is None:
                    kwn_chain_start = i
                    kwn_chain_len = 1
                else:
                    kwn_chain_len += 1
            else:
                if kwn_chain_start is not None and kwn_chain_len >= 5:
                    issues.append({
                        "type": "keep_with_next_chain",
                        "severity": "high",
                        "start_paragraph": kwn_chain_start,
                        "end_paragraph": kwn_chain_start + kwn_chain_len - 1,
                        "length": kwn_chain_len,
                        "detail": f"{kwn_chain_len} consecutive keep_with_next paragraphs (may push content to next page)",
                    })
                kwn_chain_start = None
                kwn_chain_len = 0

            # Heading style on body text
            is_heading = style_name.lower().startswith("heading") or style_name.startswith("Başlık")
            if is_heading and len(text) > 100:
                issues.append({
                    "type": "heading_on_body_text",
                    "severity": "medium",
                    "paragraph": i,
                    "style": style_name,
                    "text_length": len(text),
                    "detail": f"{style_name} on {len(text)}-char paragraph (headings should be short)",
                })

            # PageBreakBefore on non-heading
            if bool(fmt.PageBreakBefore) and not is_heading and i > 1:
                issues.append({
                    "type": "page_break_before_misuse",
                    "severity": "medium",
                    "paragraph": i,
                    "style": style_name,
                    "detail": f"PageBreakBefore on non-heading paragraph (style: {style_name})",
                })

        # Flush final chain
        if kwn_chain_start is not None and kwn_chain_len >= 5:
            issues.append({
                "type": "keep_with_next_chain",
                "severity": "high",
                "start_paragraph": kwn_chain_start,
                "end_paragraph": kwn_chain_start + kwn_chain_len - 1,
                "length": kwn_chain_len,
                "detail": f"{kwn_chain_len} consecutive keep_with_next paragraphs (may push content to next page)",
            })

        # Manual page break count via Find
        manual_breaks = 0
        try:
            find_rng = doc.Content.Duplicate
            find_rng.Find.ClearFormatting()
            while find_rng.Find.Execute(FindText="^m", Forward=True, Wrap=0):
                manual_breaks += 1
                find_rng.SetRange(find_rng.End, doc.Content.End)
        except Exception:
            pass
        if manual_breaks > 0:
            issues.append({
                "type": "manual_page_breaks",
                "severity": "info",
                "count": manual_breaks,
                "detail": f"{manual_breaks} manual page break character(s) found",
            })

        # Flag if >50% paragraphs use Heading styles
        heading_count = sum(v for k, v in style_counts.items()
                           if k.lower().startswith("heading") or k.startswith("Başlık"))
        if total_paras > 0 and heading_count > total_paras * 0.5:
            issues.append({
                "type": "excessive_heading_styles",
                "severity": "high",
                "heading_paragraphs": heading_count,
                "total_paragraphs": total_paras,
                "detail": f"{heading_count}/{total_paras} paragraphs use Heading styles (>50%)",
            })

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "total_paragraphs": total_paras,
            "issues": issues,
            "issue_count": len(issues),
            "style_summary": style_counts,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


_CORE_PROP_MAP = {
    "title": "Title",
    "subject": "Subject",
    "author": "Author",
    "keywords": "Keywords",
    "comments": "Comments",
    "category": "Category",
    "manager": "Manager",
    "company": "Company",
    "last_author": "Last Author",
}


async def word_live_set_core_properties(
    filename: str = None,
    title: str = None,
    subject: str = None,
    author: str = None,
    keywords: str = None,
    comments: str = None,
    category: str = None,
    manager: str = None,
    company: str = None,
    last_author: str = None,
) -> str:
    """[Windows only] Set Word document core/built-in properties (Title, Subject, Author, etc.).

    Equivalent to File > Info > Properties in the Word UI. Pass None for any
    field to leave it unchanged. Wrapped in undo_record so a single Ctrl+Z
    reverts every property in the call.

    Args:
        filename: Document name or path (None = active document).
        title: Document Title.
        subject: Document Subject.
        author: Author (current author / "Created by").
        keywords: Keywords (semicolon-separated by Word convention).
        comments: Free-form Comments.
        category: Category.
        manager: Manager.
        company: Company.
        last_author: "Last Author" (Last saved by).

    Returns:
        JSON {ok, document, changed: {field: {old, new}}, errors: {field: msg}}.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import (
            get_word_app, find_document, undo_record,
        )

        app = get_word_app()
        doc = find_document(app, filename)

        inputs = {
            "title": title, "subject": subject, "author": author,
            "keywords": keywords, "comments": comments, "category": category,
            "manager": manager, "company": company, "last_author": last_author,
        }

        changes: dict = {}
        errors: dict = {}

        with undo_record(app, "MCP: Set Core Properties"):
            props = doc.BuiltInDocumentProperties
            for key, value in inputs.items():
                if value is None:
                    continue
                prop_name = _CORE_PROP_MAP[key]
                try:
                    raw_old = props(prop_name).Value
                    old = str(raw_old) if raw_old is not None else None
                except Exception:
                    old = None
                try:
                    props(prop_name).Value = value
                    changes[key] = {"old": old, "new": value}
                except Exception as e:
                    errors[key] = str(e)

        return json.dumps({
            "ok": len(errors) == 0,
            "document": doc.Name,
            "changed": changes,
            "errors": errors or None,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_list_open() -> str:
    """[Windows only] List all documents currently open in Microsoft Word.

    Returns JSON with list of open documents including name, full_path,
    pages, saved status, and whether it is the active document.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_list_open
        return mac_list_open()

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app

        app = get_word_app()

        # ActiveDocument access can throw on broken/proxy state — degrade gracefully.
        try:
            active_fullname = app.ActiveDocument.FullName if app.Documents.Count > 0 else None
        except Exception:
            active_fullname = None

        try:
            count = app.Documents.Count
        except Exception as e:
            return json.dumps({
                "error": f"could not enumerate Documents collection: {e}",
                "documents": [],
            })

        documents = []
        for i in range(1, count + 1):
            entry = {"index": i}
            try:
                doc = app.Documents(i)
            except Exception as e:
                entry.update({
                    "name": "<unavailable>",
                    "error": f"could not access Documents({i}): {e}",
                })
                documents.append(entry)
                continue

            # Defensive per-property access — one broken doc must not
            # block reporting on healthy ones. Each failure is recorded
            # under entry["errors"] but does not abort the loop.
            errors = []

            def _get(attr, transform=None):
                try:
                    val = getattr(doc, attr)
                    return transform(val) if transform else val
                except Exception as ex:
                    errors.append(f"{attr}: {ex}")
                    return None

            entry["name"] = _get("Name")
            entry["full_path"] = _get("FullName")
            entry["saved"] = _get("Saved", bool)
            entry["track_revisions"] = _get("TrackRevisions", bool)

            try:
                entry["pages"] = doc.ComputeStatistics(2)  # wdStatisticPages
            except Exception:
                entry["pages"] = None

            entry["active"] = (
                active_fullname is not None
                and entry.get("full_path") == active_fullname
            )
            if errors:
                entry["errors"] = errors
            documents.append(entry)

        return json.dumps({
            "success": True,
            "count": len(documents),
            "documents": documents,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


# ---------------------------------------------------------------------------
# Field code / EndNote inspection tools
# ---------------------------------------------------------------------------

async def word_live_list_fields(
    filename: str = None,
    field_type: str = None,
    contains: str = None,
    include_code: bool = True,
    include_result: bool = True,
    max_result_chars: int = 200,
    max_code_chars: int = 500,
    include_all_stories: bool = False,
    limit: int = 200,
) -> str:
    """[Windows only] Enumerate all fields in an open Word document via Document.Fields.

    Exposes hidden field codes (w:instrText), displayed results, character offsets,
    and structured EndNote CWYW citation metadata (Record Numbers, Authors, Years).
    Works on unsaved in-memory documents without altering Word display state.

    Args:
        filename: Document name or path (None = active document).
        field_type: Filter by field type name or ID (e.g. "cite", "addin", "ref", "toc", "page").
            "cite" matches EndNote Addin citation fields and Word Citation fields.
        contains: Case-insensitive substring filter applied to field code text or result text.
        include_code: Include code text/preview in response (default True).
        include_result: Include result text preview in response (default True).
        max_result_chars: Max characters of result text to return per field (default 200).
        max_code_chars: Max characters of code text to return per field (default 500). Use word_live_get_field for full untruncated XML.
        include_all_stories: If True, also scans fields in footnotes, endnotes, and headers (default False).
        limit: Max fields to return (default 200).

    Returns:
        JSON with list of fields (index, type, type_id, char_start, char_end, code, result_preview, endnote).
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        import pythoncom
        pythoncom.CoInitialize()
        from word_document_server.core.word_com import get_word_app, find_document
        from word_document_server.utils.field_utils import get_field_type_name, parse_endnote_field_code

        app = get_word_app(filename)
        doc = find_document(app, filename)

        # Collect field objects to inspect
        field_sources = []
        if include_all_stories:
            try:
                story = doc.StoryRanges(1)
                while story is not None:
                    field_sources.append(story.Fields)
                    story = story.NextStoryRange
            except Exception:
                field_sources = [doc.Fields]
        else:
            field_sources = [doc.Fields]

        fields = []
        total_fields_count = 0
        current_index = 0

        ft_filter = field_type.strip().lower() if field_type else None
        contains_filter = contains.strip().lower() if contains else None

        for fld_collection in field_sources:
            try:
                col_count = fld_collection.Count
                total_fields_count += col_count
            except Exception:
                continue

            for i in range(1, col_count + 1):
                current_index += 1
                try:
                    f = fld_collection(i)
                except Exception:
                    continue

                # Safely extract properties
                try:
                    type_id = int(f.Type)
                except Exception:
                    type_id = -1
                type_name = get_field_type_name(type_id)

                code_text = ""
                try:
                    code_text = str(f.Code.Text) if f.Code and f.Code.Text else ""
                except Exception as ex:
                    code_text = f"<error reading code: {ex}>"

                result_text = ""
                try:
                    result_text = str(f.Result.Text) if f.Result and f.Result.Text else ""
                except Exception as ex:
                    result_text = f"<error reading result: {ex}>"

                char_start = -1
                char_end = -1
                try:
                    if f.Result:
                        char_start = f.Result.Start
                        char_end = f.Result.End
                except Exception:
                    pass
                if char_start < 0:
                    try:
                        if f.Code:
                            char_start = f.Code.Start
                            char_end = f.Code.End
                    except Exception:
                        pass

                is_locked = False
                try:
                    is_locked = bool(f.Locked)
                except Exception:
                    pass

                endnote_data = parse_endnote_field_code(code_text)

                # Filter by field_type if specified
                if ft_filter:
                    matched_type = False
                    if ft_filter in ("cite", "citation", "endnote"):
                        if (endnote_data is not None and endnote_data.get("field_category") == "citation") or type_id == 1982 or "cite" in code_text.lower():
                            matched_type = True
                    elif ft_filter in ("bib", "bibliography", "reflist"):
                        if (endnote_data is not None and endnote_data.get("field_category") == "bibliography") or type_id == 1983 or "reflist" in code_text.lower():
                            matched_type = True
                    elif ft_filter == "addin":
                        if type_id == 26 or "addin" in type_name.lower():
                            matched_type = True
                    elif ft_filter in ("toc", "table of contents"):
                        if type_id == 13 or "toc" in type_name.lower():
                            matched_type = True
                    elif ft_filter in ("page", "numpages"):
                        if type_id in (32, 33) or "page" in type_name.lower():
                            matched_type = True
                    elif ft_filter in ("ref", "cross-reference"):
                        if type_id == 3 or "ref" in type_name.lower():
                            matched_type = True
                    elif ft_filter in ("hyperlink", "link"):
                        if type_id == 88 or "hyperlink" in type_name.lower():
                            matched_type = True
                    else:
                        if ft_filter == str(type_id) or ft_filter in type_name.lower() or ft_filter in code_text.lower():
                            matched_type = True

                    if not matched_type:
                        continue

                # Filter by contains if specified
                if contains_filter:
                    if contains_filter not in code_text.lower() and contains_filter not in result_text.lower():
                        continue

                field_entry = {
                    "index": current_index,
                    "type": type_name,
                    "type_id": type_id,
                    "char_start": char_start,
                    "char_end": char_end,
                    "is_locked": is_locked,
                }

                if include_code:
                    if len(code_text) > max_code_chars:
                        field_entry["code"] = code_text[:max_code_chars] + f" ...[truncated {len(code_text) - max_code_chars} chars, use word_live_get_field]"
                        field_entry["code_truncated"] = True
                    else:
                        field_entry["code"] = code_text
                        field_entry["code_truncated"] = False

                if include_result:
                    if len(result_text) > max_result_chars:
                        field_entry["result_preview"] = result_text[:max_result_chars] + " ...[truncated]"
                    else:
                        field_entry["result_preview"] = result_text

                if endnote_data:
                    field_entry["endnote"] = endnote_data

                fields.append(field_entry)
                if len(fields) >= limit:
                    break

            if len(fields) >= limit:
                break

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "total_fields": total_fields_count,
            "matched_fields": len(fields),
            "fields": fields,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_field(
    filename: str = None,
    field_index: int = None,
    char_position: int = None,
) -> str:
    """[Windows only] Get a single field by index or character position, returning untruncated code and result.

    Ideal for retrieving full EndNote post-annotation XML payloads (5-10 KB) without
    truncation or loading the entire document.xml.

    Args:
        filename: Document name or path (None = active document).
        field_index: 1-based index in Document.Fields.
        char_position: Document character offset. Locates the field enclosing or nearest to this position.

    Returns:
        JSON with untruncated code, result, type, offsets, and complete parsed EndNote metadata.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if field_index is None and char_position is None:
        return json.dumps({"error": "Either field_index or char_position must be specified"})

    try:
        import pythoncom
        pythoncom.CoInitialize()
        from word_document_server.core.word_com import get_word_app, find_document
        from word_document_server.utils.field_utils import get_field_type_name, parse_endnote_field_code

        app = get_word_app(filename)
        doc = find_document(app, filename)

        total_fields = doc.Fields.Count
        if total_fields == 0:
            return json.dumps({"error": f"Document '{doc.Name}' contains no fields"})

        target_field = None
        target_index = None

        if field_index is not None:
            if field_index < 1 or field_index > total_fields:
                return json.dumps({
                    "error": f"field_index {field_index} is out of range (document has {total_fields} fields, 1-indexed)"
                })
            target_field = doc.Fields(field_index)
            target_index = field_index
        else:
            # Find field enclosing or closest to char_position
            closest_dist = float("inf")
            for i in range(1, total_fields + 1):
                f = doc.Fields(i)
                f_start = f.Result.Start if f.Result else (f.Code.Start if f.Code else -1)
                f_end = f.Result.End if f.Result else (f.Code.End if f.Code else -1)
                if f_start <= char_position <= f_end:
                    target_field = f
                    target_index = i
                    break
                dist = min(abs(char_position - f_start), abs(char_position - f_end))
                if dist < closest_dist:
                    closest_dist = dist
                    target_field = f
                    target_index = i

        if target_field is None:
            return json.dumps({"error": "Field not found"})

        try:
            type_id = int(target_field.Type)
        except Exception:
            type_id = -1
        type_name = get_field_type_name(type_id)

        code_text = ""
        try:
            code_text = str(target_field.Code.Text) if target_field.Code and target_field.Code.Text else ""
        except Exception as ex:
            code_text = f"<error reading code: {ex}>"

        result_text = ""
        try:
            result_text = str(target_field.Result.Text) if target_field.Result and target_field.Result.Text else ""
        except Exception as ex:
            result_text = f"<error reading result: {ex}>"

        char_start = -1
        char_end = -1
        try:
            if target_field.Result:
                char_start = target_field.Result.Start
                char_end = target_field.Result.End
        except Exception:
            pass
        if char_start < 0:
            try:
                if target_field.Code:
                    char_start = target_field.Code.Start
                    char_end = target_field.Code.End
            except Exception:
                pass

        is_locked = False
        try:
            is_locked = bool(target_field.Locked)
        except Exception:
            pass

        endnote_data = parse_endnote_field_code(code_text)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "field": {
                "index": target_index,
                "type": type_name,
                "type_id": type_id,
                "char_start": char_start,
                "char_end": char_end,
                "is_locked": is_locked,
                "code": code_text,
                "code_length": len(code_text),
                "result": result_text,
                "result_length": len(result_text),
                "endnote": endnote_data,
            }
        }, ensure_ascii=False)

    except Exception as e:
        import traceback
        return json.dumps({
            "error": str(e),
            "traceback": traceback.format_exc(),
        })


# ---------------------------------------------------------------------------
# Macro discovery tools
# ---------------------------------------------------------------------------

async def word_live_list_macros(
    filename: str = None,
    filter: str = None,
    source: str = "all",
    limit: int = 100,
) -> str:
    """[Windows only] List callable VBA macros and add-in commands in the active document, templates, and add-ins.

    Discovers macros from:
    - Active document (if macro-enabled .docm)
    - Attached document template and global Normal.dotm
    - Word STARTUP templates and installed Add-Ins (e.g. EndNote CWYW, Zotero)

    Enables agents to discover exact macro names (such as 'ENFormatBibliography' to update
    citations/bibliography, or 'ZoteroRefresh') before executing them with word_live_run_macro.
    Works reliably even when 'Trust access to the VBA project object model' is disabled in Office Trust Center.

    Args:
        filename: Document name or path (None = active document).
        filter: Optional case-insensitive substring search matching macro name, module, or template (e.g. "biblio", "citation", "endnote", "zotero").
        source: Filter by origin: "all" (default), "document", "templates", or "addins".
        limit: Maximum number of macros to return (default 100).

    Returns:
        JSON with list of callable macros (name, module, qualified_name, source, source_type, path).
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        import pythoncom
        pythoncom.CoInitialize()
        from word_document_server.core.word_com import get_word_app, find_document
        from word_document_server.utils.macro_utils import list_available_macros

        app = None
        doc = None
        try:
            app = get_word_app()
            doc = find_document(app, filename)
        except Exception:
            app = None
            doc = None

        macros = list_available_macros(app, doc, query=filter, source_filter=source)

        return json.dumps({
            "success": True,
            "document": doc.Name if doc else None,
            "total_macros": len(macros),
            "matched_macros": min(len(macros), limit),
            "macros": macros[:limit],
        }, ensure_ascii=False)

    except Exception as e:
        import traceback
        return json.dumps({
            "error": str(e),
            "traceback": traceback.format_exc(),
        })
