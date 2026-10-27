# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.7.0] - 2026-10-26

### Added
- **Stable Comment IDs & Target Echo (Phase 1)**:
  - Generates deterministic 12-hex IDs (`comment_id`, e.g. `c_a1b2c3d4e5f6`) based on creation timestamp, author, and range start in `word_live_get_comments` and `word_live_add_comment`.
  - Added `comment_id` resolution support to `word_live_reply_to_comment`, `word_live_resolve_comment`, and `word_live_delete_comment`, with backward-compatible fallback to `comment_index`.
  - Added operation echo payload returning target comment details (`replied_to`, `resolved_comment`, `deleted_comment` with author and text excerpt) and index shift warnings to prevent off-by-one errors after deletions or additions.
- **Docstring Conventions & Invariant Contracts (Phase 2)**:
  - Added standardized `CONVENTIONS & INVARIANTS` sections across `word_live_get_text`, `word_live_find_text`, `word_live_replace_text`, `word_live_insert_text`, `word_live_format_text`, `word_live_delete_text`, `word_live_insert_paragraphs`, and comment tools.
  - Documents 0-based character offsets vs 1-based paragraph indices, offset invalidation behavior during edits, and strict tracked changes permission policies.
- **Figure Inventory & In-Place Replacement (Phase 4)**:
  - Added `word_live_list_inline_shapes` tool returning complete figure inventory: dimensions (pt/inches), type classification (`inline_picture`, `linked_picture`, `chart`, `embedded_ole`), link target paths, paragraph indices, and page numbers.
  - Added `word_live_replace_inline_image` tool supporting in-place figure swapping with automatic dimension preservation (`keep_size=True`), linked vs embedded picture handling, undo grouping, and tracked revision logging.
- **Find Coordinate Bridge & Span Deduplication (Phase 5)**:
  - Enhanced `word_live_find_text` matches to include 1-based `paragraph_index` (derived via `doc.Range(0, start).Paragraphs.Count`) and `page` number (`wdActiveEndAdjustedPageNumber`).
  - Added `(start, end)` span deduplication in `word_live_find_text` to prevent duplicate match reports from overlapping ranges or repeated COM hits.
- **Per-Occurrence & Offset-Bounded Replacement (Phase 6)**:
  - Extended `word_live_replace_text` with `occurrence: int` (1-based instance target) and `start_after: int` (0-based starting offset).
  - Enforces `replace_all=False` when targeting a specific occurrence, records exact replaced character spans, and raises descriptive errors if the requested occurrence exceeds match counts.
- **`replace_all` / `track_changes` Smart Defaults**:
  - `word_live_replace_text` now smartly defaults `replace_all = False` when `track_changes=True` or when `start_after` is specified, eliminating agent friction and preventing infinite loop guard failures while maintaining safe single-replacement behavior.
- **Word COM STA Serialization & Auto-Retry**:
  - Implemented `@serialize_word_com` decorator powered by `AsyncRLock` across all 51 live Word COM tools (`word_live_*`), automatically serializing concurrent async calls across Word's Single-Threaded Apartment (STA) and retrying transient `RPC_E_CALL_REJECTED` (`0x80010001` / `-2147418111`) errors with exponential backoff.
- **Multi-Document Target Warning & Focus Safety**:
  - Added `get_document_target_warning` and stderr warnings when multiple documents are open in Word and `filename` is omitted. Responses now return a `"target_warning"` field identifying the targeted frontmost document and advising callers to specify `filename` explicitly to prevent mis-targeting when active focus changes.

### Fixed
- **D0 (Regression)**: Fixed `TypeError: word_live_replace_text() got multiple values for argument 'occurrence'` in server tool adapter by passing all arguments via keyword.
- **D2 (High)**: Fixed `word_live_replace_inline_image` duplicating figures instead of replacing in-place by calling `old_shape.Delete()` after `AddPicture`, ensuring paired tracked deletion/insertion revisions, correct paragraph positioning, and added `force_embed` parameter.
- **D1, D3, D6 (High/Medium/Low)**: Resolved comment ID collision and reply double-counting by filtering out threaded replies (`c.Ancestor is not None`) from top-level comment listings in `word_live_get_comments`, formatting reply IDs deterministically as `<parent_id>_r<index>`, and returning distinct `reply_comment_id` in `word_live_reply_to_comment`.
- **D4 (Medium)**: Fixed raw COM tuple error when deleting comments with replies by executing `DeleteRecursively()` when replies exist, exposing `delete_recursively: bool = True`, and cleanly catching COM errors.
- **D5 (Medium-High)**: Fixed off-by-one paragraph index attribution in `word_live_find_text`, `word_live_list_inline_shapes`, and `word_live_replace_inline_image` when a match begins on a paragraph's initial character by evaluating `doc.Range(0, min(doc_end, start + 1)).Paragraphs.Count`.
- **D7 (Low)**: Added explicit index base definitions and field-code offset inflation documentation to `word_live_get_text` and `word_live_get_page_text`.

## [1.6.22] - 2026-10-26

### Added
- **Arbitrary length search and replacement in `word_live_find_text` and `word_live_replace_text`** — transparently bypasses Microsoft Word's 255-character `Find.Execute` COM parameter limit via chunked prefix matching and direct Word `Range` text assignment. Search strings can now be 500, 1,000, or more characters (e.g. full paragraphs or multi-sentence blocks).
- New utility module `word_document_server/utils/find_utils.py` — handles safe prefix chunking, newline normalization (`\r\n` vs Word's `\r`), string verification, and word boundary checks.
- Comprehensive test suite in `tests/test_find_utils.py` and `tests/test_live_find_replace_long.py` (29 new tests) covering edge cases (exact 255 chars, 256 chars, 500 chars, 1,000+ chars, false-positive prefix rejection, case sensitivity, whole word boundaries, deletions, replacements across paragraph newlines).
- **Tracked changes requirement for modification tools** — updated docstrings and schemas for all live modification tools (`word_live_replace_text`, `word_live_insert_text`, `word_live_format_text`, `word_live_delete_text`, `word_live_modify_table`, `word_live_add_table`, `word_live_apply_list`, `word_live_insert_paragraphs`) explicitly instructing agents to use `track_changes=True` unless granted permission for untracked edits.

### Fixed
- **`word_live_find_text` duplicate match bug** — previously advanced search via `rng.SetRange(match_end, doc.Content.End)` which caused Word's `Find` object to lose its search pointer and re-match the exact same position up to `max_results` (50) times. Now cleanly advances the collapsed insertion point via `rng.SetRange(adv_pos, adv_pos)` with position-regression guards.

## [1.6.0] - 2026-04-29

### Added
- **`word_live_set_core_properties`** — set Word document Title, Subject, Author, Keywords, Comments, Category, Manager, Company, Last Author via `Document.BuiltInDocumentProperties`. Wrapped in `undo_record` so a single Ctrl+Z reverts every property in one call. Equivalent to File > Info > Properties in the Word UI.
- New `word_document_server/utils/text_safety.py` — shared `reject_control_chars()` validator for Find/Replace/Insert text inputs.
- `scrub_orphans` parameter on `word_live_modify_table` `delete_table` operation (default `True`) — cleans stranded `\x07` cell-separator bytes the Word COM `Table.Delete()` occasionally leaves behind.

### Fixed
- **`word_live_replace_text` data-loss vector** — passing `\x07` (cell separator) as `find_text` previously matched across cell boundaries and could delete entire documents. Control bytes (U+0000–U+001F except `\t`, `\n`, `\r`) now rejected with a descriptive error before Find.Execute is reached.
- **`word_live_find_text`** — same control-byte protection applied to `search_text`.
- **`word_live_insert_text`** — same control-byte protection applied to `text` (prevents inserting orphan cell separators outside a real table).
- **`word_live_modify_table` `delete_table`** — leftover `\x07` separators after Word's native `Table.Delete()` now scrubbed by default (configurable via `scrub_orphans=False`).
- **`word_live_add_table`** — rejects `position` offset that falls inside an existing table's range or sits immediately after an orphan cell separator (would otherwise silently merge new content into existing/residual table structure).
- **`word_live_setup_heading_numbering`** — paragraphs that previously kept a custom template style (e.g. `Font Style30/31`) after a forced heading reassignment now (a) get explicit per-paragraph style assignment with try/except, (b) receive the same font/size/bold/color customizations as direct formatting so visual output matches even when the underlying style refuses to swap, (c) report any failed reassignments under a new `restyle_failures` field in the response.
- **`word_live_modify_table`** — re-reads `Tables.Count` per call and validates `table_index` with a helpful message ("table_index N out of range. Document has K table(s)…") instead of throwing "Document has no tables" when a stale index is passed after a prior delete.
- **`word_live_list_open`** — defensive per-document property access; one document in a broken COM proxy state no longer aborts the whole call. Each document entry now includes `index`, `track_revisions`, and per-property `errors` array.
- **`word_live_find_text`** — defensive `Range.Text` / `Range.Start` / `Range.End` / `Document.Name` access via internal `_safe_attr` helper. Transient COM marshalling failures after MCP reconnect now produce partial matches with `<unreadable>` placeholders and a `partial_errors` array, rather than aborting the call.

## [1.5.1] - 2026-04-08

### Fixed
- `word_live_replace_text` — infinite loop when wildcard pattern matches zero-length strings (e.g., `*` alone); now skips forward on zero-length matches and enforces 50K replacement safety ceiling

## [1.5.0] - 2026-04-08

### Added
- **macOS live editing support** via JavaScript for Automation (JXA) — 33 of 41 `word_live_*` tools now work on macOS with Word for Mac
- New module `word_document_server/core/word_mac.py` — JXA bridge with 30+ functions for Word for Mac automation
- Platform auto-detection: same tool names and parameters on both Windows and macOS
- `pywin32` as conditional dependency (Windows only) in `pyproject.toml`

### Changed
- All `print()` calls in `main.py` redirected to stderr — fixes MCP stdio protocol corruption that prevented the server from loading in some clients
- All live tool functions now dispatch to macOS JXA implementations when `sys.platform == "darwin"`
- Updated tool count: 76 cross-platform + 41 Windows Live + 33 macOS Live

### Not Available on macOS
These 4 tools require Windows COM APIs with no AppleScript/JXA equivalent:
- `word_live_get_undo_history` — undo stack inspection not exposed in Word for Mac's scripting dictionary
- `word_live_reply_to_comment` — threaded comment replies not in AppleScript dictionary
- `word_live_resolve_comment` — comment Done property not in AppleScript dictionary
- `word_live_add_watermark` — requires VBA `Shapes.AddTextEffect` (VBA bridge killed by Apple sandboxing in Word 365)

## [1.4.1] - 2026-04-08

### Fixed
- `word_live_replace_text` — `^s` (non-breaking space) now converted to `\u00a0` in replacement text (#4)

## [1.4.0] - 2026-04-08

### Added
- `word_live_insert_paragraphs` — insert multiple paragraphs near a target (by text or index) in a single undo record
- `word_live_take_snapshot` — store paragraph baseline for efficient change detection
- `word_live_get_diff` — compare current document against snapshot, returns only changed paragraphs
- `word_live_snapshot_status` — check snapshot existence and age
- `word_live_modify_table` — new `set_row` and `set_range` operations for bulk cell updates

### Fixed
- `word_live_replace_text` — infinite loop when document has TrackRevisions enabled independently of `track_changes` parameter (#7)
- All destructive tools now unconditionally restore `doc.TrackRevisions` in `finally` block

### Credits
- Snapshot/diff tools, `insert_paragraphs`, and bulk table operations adapted from PR #5 by @FarhadGSRX

## [1.3.0] - 2026-02-28

### Added
- `word_live_modify_table` — table operations via COM: get info, set cell, add/delete rows/columns, merge cells, autofit, delete table
- `word_live_save` — save document in place or save-as (docx, pdf, rtf, txt)
- `word_live_toggle_track_changes` — toggle or explicitly set track changes mode on/off
- `word_live_insert_image` — insert image with sizing, alignment, wrapping, and optional border
- `word_live_insert_cross_reference` — insert live cross-references to headings, bookmarks, figures, tables, equations, footnotes, endnotes
- `word_live_list_cross_reference_items` — list available cross-reference targets with their indices
- `word_live_insert_equation` — insert mathematical equations using UnicodeMath syntax
- `word_live_reply_to_comment` — threaded comment replies (Word 2016+)
- `word_live_resolve_comment` — mark comments as resolved/unresolved (Word 2016+)
- `word_live_delete_comment` — permanently delete a comment
- Total tool count now **114** (75 cross-platform + 39 Windows Live)

### Changed
- `word_live_delete_text` — now table-aware: deletes table objects within range before text deletion
- `word_live_insert_text` — auto-chunks text >30K chars to avoid COM 32K limit
- `word_live_setup_heading_numbering` — handles inflated paragraph ranges from comment anchors
- `word_live_modify_table` set_cell operation now accepts tracked changes before writing to prevent layered content

## [1.2.0] - 2025-02-15

### Added
- `word_live_replace_text` — find & replace via COM that works across tracked change boundaries; supports wildcards (`^m`, `^t`, `^p`) and tracked changes mode
- `word_live_diagnose_layout` — read-only scan for layout problems: keep_with_next chains, heading styles on body text, PageBreakBefore misuse, manual breaks
- `word_live_get_paragraph_format` — inspect paragraph formatting (font, spacing, alignment, list info, style); `include_runs=True` for per-run detail
- `word_live_get_page_text` — read text from specific page(s) with char offsets for chaining into format/edit tools
- `word_live_get_undo_history` — list undo stack entries
- `word_live_apply_list` — apply bullet, numbered, or multilevel list formatting
- `word_live_setup_heading_numbering` — auto-numbered headings (1. / 1.1) via multilevel list linked to Heading styles; configurable style params (font, size, color, spacing)

### Changed
- `word_live_format_text` — added `paragraph_alignment`, `page_break_before`, paragraph-index addressing (`start_paragraph`/`end_paragraph`), `preserve_direct_formatting` for style changes
- `word_live_find_text` — added `use_wildcards` for `^m`/`^t`/`^p`/Word wildcard syntax; `context_chars` now configurable (default 60, was 30)
- `word_live_set_paragraph_spacing` — clarified that `line_spacing` is in points (1.15 lines = 13.8pt)

## [1.1.0] - 2025-01-10

### Added
- 27 Windows Live tools (`word_live_*`) using COM automation for editing documents open in Word
- Per-operation undo system — all destructive tools wrapped with `UndoRecord`; each tool call = one Ctrl+Z entry
- `word_live_undo` — programmatic undo of last N operations
- Live editing tools: `word_live_insert_text`, `word_live_delete_text`, `word_live_format_text`, `word_live_add_table`
- Live reading tools: `word_live_get_text`, `word_live_get_info`, `word_live_find_text`
- Live comment & revision tools: `word_live_add_comment`, `word_live_get_comments`, `word_live_list_revisions`, `word_live_accept_revisions`, `word_live_reject_revisions`
- Live layout tools: `word_live_set_page_layout`, `word_live_add_header_footer`, `word_live_add_page_numbers`, `word_live_add_section_break`, `word_live_set_paragraph_spacing`, `word_live_add_bookmark`, `word_live_add_watermark`
- `word_screen_capture` — screenshot of the Word window
- Cross-platform tracked changes: `track_replace`, `track_insert`, `track_delete`, `list_tracked_changes`, `accept_tracked_changes`, `reject_tracked_changes`
- Cross-platform comments: `add_comment` anchored to text
- Cross-platform hyperlinks: `manage_hyperlinks` (add, list, remove, update)
- Cross-platform layout tools: `set_page_layout`, `add_header_footer`, `add_page_numbers`, `add_section_break`, `set_paragraph_spacing`, `add_bookmark`, `add_watermark`
- Cross-platform footnote tools (10): add, delete, validate, customize footnotes and endnotes
- Cross-platform protection tools: `protect_document`, `unprotect_document`, `add_restricted_editing`, `add_digital_signature`, `verify_document`
- Multiple transport support: stdio (default), SSE, streamable-http
- `MCP_AUTHOR` / `MCP_AUTHOR_INITIALS` environment variables for author metadata
- PyPI packaging as `word-mcp-live`

## [1.0.0] - 2024-12-01

### Added
- Initial release based on [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)
- 54 cross-platform tools using python-docx
- Document management, content editing, formatting, tables, extraction
- FastMCP server with stdio transport

[1.5.1]: https://github.com/ykarapazar/word-mcp-live/compare/v1.5.0...v1.5.1
[1.5.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.4.1...v1.5.0
[1.4.1]: https://github.com/ykarapazar/word-mcp-live/compare/v1.4.0...v1.4.1
[1.3.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/ykarapazar/word-mcp-live/releases/tag/v1.0.0
