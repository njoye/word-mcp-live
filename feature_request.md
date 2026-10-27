Feature Request: Comment Robustness, Shape Replacement, and Find/Replace Precision for the Word MCP (live/COM tools)
Requested tool family: word_live_* comment, shape, and find/replace tools
Priority: High — eliminates the dominant agent-error class (wrong-thread replies) and the last manual-work gaps observed in a production manuscript-editing session (2026-10-19 → 10-23, 944 messages, ~170 tool calls)
Date: 2026-10-23
Excluded (in flight separately): word_live_update_citations (CWYW refresh trigger)
1. Problem Statement
These five requests come from one long production session editing an EndNote-managed manuscript with tracked changes and 30+ review comments. Each item caused measurable wasted effort: wrong-thread comments requiring multi-step cleanup, tool calls that pushed the agent into untracked placeholder-token hacks, and one workflow (figure swap) that had to be handed back to the human because no tool existed.
Evidence per feature is quoted or referenced below so acceptance criteria can be validated against real failure modes.

2. Feature F1 — Stable comment IDs + target echo (highest priority)
2.1 Observed failures
word_live_reply_to_comment / word_live_delete_comment address comments by positional index, and every write mutates the index space:

A reply is inserted as a new top-level entry immediately after the parent → all later indices shift +1.
A delete shifts all later indices −1.
User edits in the Word UI (deleting/resolving a comment thread) invalidate the agent's snapshot silently.

Five confirmed wrong-target incidents in one session:

Agent replied to its own reply — after replies at 19→(20) and 21→(22), it replied to "22", which had become its own COCO reply; the 640-justification text nested under the wrong node. Caught only by manual index reconstruction: "I replied to my own reply! Crap."
640-justification reply attached to the COCO thread — delete + re-post + re-fetch cycle.
Throttling reply attached to the "$209 is cut off" thread — undetected for two days, discovered only incidentally.
YOLO26 reply landed on two wrong threads ("not fully real" AND AI HAT+) — two delete/re-post/verify cycles.
comment_index 36 out of range (1-33) — the user deleted a thread in the Word UI between the agent's get_comments and its write.

The agent derived the correct mitigation ("process in descending index order") twice in thinking blocks and still failed between calls, because the invariant lived in the model's scratch reasoning instead of the tool contract. Each recovery cost a full get_comments fetch (39 comments, tens of KB of context).
2.2 Proposed API changes
a) Stable IDs. word_live_get_comments returns, per comment and per reply:

        
          
            json
            
              Copy
            
          
          {
  "comment_id": "c_638264832000000000",   // derived from Word Comment.Timestamp (immutable)
  "index": 27,                             // kept for backward compat
  ...
}
        
      word_live_reply_to_comment, word_live_delete_comment, word_live_resolve_comment, word_live_get_comments_by_* accept comment_id instead of comment_index (keep index as a deprecated fallback). ID resolution happens server-side at call time → index arithmetic never leaves the server.
b) Target echo on every comment write. All comment-mutating responses include:

        
          
            json
            
              Copy
            
          
          {
  "success": true,
  "action": "reply",
  "replied_to": { "comment_id": "c_638264832...", "author": "Timo Mueller", "excerpt": "Probably need a citation ehre" },
  "reply_index": 28,
  "index_shift_note": "comments after index 28 shifted +1"
}
        
      A wrong-thread hit becomes visible in the tool result, without re-fetching the whole list. This alone would have caught all five incidents at the call that caused them.
c) Docstring rule (see also F2): "Replies and deletions shift positional indices. UI-side edits invalidate indices. Prefer comment_id."
2.3 Implementation sketch
Word's Comment.Index is positional, but Comment.Timestamp is stable per comment. Maintain the mapping inside each call: enumerate doc.Comments, key by timestamp, resolve incoming comment_id. No persistence needed across calls. ~20 lines plus response-shape changes.
2.4 Acceptance criteria

get_comments returns unique comment_id per comment and reply; IDs survive replies/deletes on other threads.
Reply by comment_id lands on the correct thread even when 3 replies were inserted above it since the last fetch.
Every reply/delete/resolve response contains the acted-on thread's excerpt + author.
A stale comment_id (comment deleted in UI) returns a clear error naming the missing ID — never a silent fallback to a positional guess.


3. Feature F2 — Conventions block in tool docstrings (zero code, high value)
3.1 Observed failures
The agent repeatedly rediscovered undocumented behavior at runtime cost:
BehaviorHow it was discoveredCostReply inserts shift later indices +1Empirically, after wrong-thread incidents~50 messages, 5 cleanup cyclesUnder Track Changes, deleted text stays in the model → char offsets do NOT shift after tracked deletesReasoned out mid-COM-script for the Garcia-Perez edit ("track changes preserves positions")Risk-laden live scriptingfind_text 255-char Find limitSilent 0-match results, then diagnosisSeveral failed searchesMixed 0-/1-based paragraph indexing across tools (file tools 0-based python-docx; live tools 1-indexed; insert_paragraphs target 0-based)Trial and error / near-missesRepeated verification fetchesreplace_text tracked vs. untracked semanticsGuessed; agent avoided tracked mode for placeholder swaps (violating its own plan rule)Untracked edits in a tracked workflow
3.2 Proposal
Add a shared conventions paragraph to every word_live_* docstring (or an mcp://conventions resource if the server supports it):

Index bases: state 0- or 1-based explicitly per tool; state which coordinate system each tool's outputs feed (char offsets vs. paragraph indices).
Comment index-shift rule (from F1c).
Tracked-change offset stability: "With TrackRevisions on, deletions are marked, not removed; document character offsets remain valid until revisions are accepted."
Find limits: 255-char search string limit; wildcard escapes (^m, ^t, ^p, ^s).
Anchor guidance (validated in-session): "Use at least one full, unique sentence as the find anchor; check match_count == 1 before replacing. Short anchors ('crossings', 'YOLO26') hit multiple locations." Every full-sentence-anchored replacement in the session landed first-try; every short anchor cost disambiguation round-trips.

3.3 Acceptance criteria
Docstrings for find_text, replace_text, insert_text, delete_text, format_text, get_page_text, get_text, all comment tools each state their index base and the relevant rules; no rule appears only in external docs.

4. Feature F4 — Inline shape inventory + in-place image replacement
4.1 Observed failure
Regenerated FIG. 1–4 PNGs could not be swapped into the document via MCP; the step was handed back to the human. The agent's analysis of its options (verbatim from session): insert-new + delete-one-graphic-character was "risky"; it could not determine embedded vs. linked pictures without raw COM; delete_text on a graphic character under tracked changes was untested territory. It ultimately verified figures by parsing SVG coordinates instead — a workaround, not the workflow.
4.2 Proposed API
word_live_list_inline_shapes(filename, page=null, limit=100) → per shape:

        
          
            json
            
              Copy
            
          
          {
  "shape_index": 3,            // 1-based, InlineShapes order
  "name": "Picture 4",
  "type": "inline_picture",    // inline_picture | linked_picture | shape | chart | equation ...
  "paragraph_index": 118,      // 1-based, same convention as other live tools
  "char_start": 61847, "char_end": 61848,
  "width_pt": 415.0, "height_pt": 290.5,
  "linked": false, "source_file": null,   // link target if linked
  "alt_text": "FIG. 1 ..."
}
        
      word_live_replace_inline_image(filename, shape_index, image_path, keep_size=true, track_changes=true)

Replaces the picture's InlineShapes.Item(i) content in place (Range.InlineShapes.AddPicture(..., LinkToFile:=false, SaveWithDocument:=true) inserted at the shape's range, old shape deleted), preserving width/height and paragraph position.
keep_size=false → fit-to-column.
Under Track Changes: old image recorded as deletion, new as insertion (consistent with the rest of the toolset).
If the target shape is a linked picture, prefer LinkFormat.Update() when image_path is null; error if path given but shape is linked and force_embed not set.

4.3 Acceptance criteria

Document with 4 embedded PNGs + 1 linked PNG: list_inline_shapes returns 5 entries, correctly flagging the linked one with its source path.
replace_inline_image on an embedded shape swaps the image, preserves display size and paragraph position; one Ctrl+Z fully reverts.
Replacing a linked shape without force_embed updates the link instead of embedding.
With Track Changes on, the swap appears as delete+insert revisions, not a silent overwrite.


5. Feature F5 — find_text: dedupe matches + paragraph cross-reference
5.1 Observed failures

Duplicate matches. The same position returned up to 5× with identical start/end (the "match_count: 50" Zhang search was ~10 unique hits repeated 5×; the XUMARK case: "the duplicate matches are just quirks of the search tool"). The agent mis-planned replacement counts twice because match_count was untrustworthy.
No coordinate bridge. find_text/get_page_text return char offsets; format_text, insert_paragraphs, add_picture take paragraph indices. Every find→edit pipeline required manual translation or a redundant get_page_text call.

5.2 Proposal

Dedupe returned matches by (start, end) server-side before counting; match_count = unique positions. (Likely a side effect of Find loop re-entering zero-width/field-adjacent ranges — dedupe is the safe contract either way.)
Add "paragraph_index": N (1-based, live-tool convention) and "page": N to each match, derived from Range.Paragraphs(1) / Information(wdActiveEndAdjustedPageNumber).

5.3 Acceptance criteria

Searching a string that appears once returns match_count: 1 (currently returns up to 5).
Searching "Zhang, Q." in the session document returns ≈10 entries, not 50.
Each match carries a paragraph_index that works directly in word_live_format_text(start_paragraph=...).


6. Feature F6 — Per-occurrence replace_text
6.1 Observed failure
Three identical plain-text strings (Alqahtani et al. 2025) needed to become three different temporary citations (#57 twice, #51 for Xu — distinct targets per occurrence). With only replace_all or first-match semantics, the agent improvised:

        
          
            text
            
              Copy
            
          
          replace_all "{Alqahtani, 2025}" → "ALQMARK57"   (3 hits, untracked)
replace "ALQMARK57" → "{Alqahtani, 2025, #57}"   ×3 sequential untracked calls
        
      This left placeholder tokens inside the manuscript in the intermediate state and forced tracked: false on every call — silently violating the session's own "ALL edits tracked" rule — because layered tracked replacements of the same span were unreliable.
6.2 Proposal
Add to word_live_replace_text:

occurrence: int|null — replace only the Nth unique match (1-based, in document order). Mutually exclusive with replace_all=true. Error if N > match count, reporting the actual count.
start_after: int|null — begin searching after this char offset (pairs naturally with find_text results for surgical, position-pinned edits).
Combine freely with track_changes=true — each occurrence edit is its own tracked revision.

Typical agent flow becomes: find_text (3 unique matches, each with char offsets) → 3 calls with start_after=<prev_end> or occurrence=1/2/3, all tracked, no placeholder tokens ever entering the document.
6.3 Acceptance criteria

Three identical matches + occurrence=2 → exactly the middle one replaced; replacements: 1; response includes the replaced position.
occurrence=4 when 3 exist → error "3 matches found, occurrence 4 requested", zero changes.
All occurrence replacements with track_changes=true appear as independent tracked insertions (verifiable via list_tracked_changes).
start_after respects document order after prior tracked edits (offsets stable per F2 rule 3).


7. Cross-cutting notes

F1 and F2 are the same lesson: the session's costliest errors (5 wrong-thread replies) came from invariants living in the model's head instead of the tool contract. F1 makes them structurally impossible; F2 makes the remaining ones discoverable before the first failed call.
F5/F6 share a root: find_text output isn't trustworthy/expressive enough to plan precise edits, which is what pushed the agent into the F6 placeholder hack. Fixing both retires that pattern.
F4 is the only remaining "hand work back to the human" gap in the document-editing loop (citation refresh being handled separately).

8. Non-goals

No image viewing capability (explicitly out of scope).
No citation formatting or CWYW field insertion (stays with EndNote; refresh trigger is the separate in-flight request).
No change to file-based (python-docx) tools — all requests target the live/COM family.