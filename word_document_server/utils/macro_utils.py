"""Utilities for discovering and inspecting VBA macros in Microsoft Word documents and templates.

Provides parsing for OpenXML macro-enabled files (.dotm, .docm) to enumerate available
procedures without requiring 'Trust access to the VBA project object model' in Word's Trust Center.
"""

import glob
import io
import os
import re
import sys
import zipfile
from typing import Any, Dict, List, Optional, Tuple

try:
    import olefile
except ImportError:
    olefile = None


def decompress_ovba_chunk(data: bytes) -> bytes:
    """Decompress an MS-OVBA compressed stream chunk according to [MS-OVBA] 2.4.1.

    Args:
        data: Raw compressed byte sequence from a VBA module stream.

    Returns:
        Decompressed byte sequence.
    """
    if not data or data[0] != 1:
        return b""
    pos = 1
    decompressed = bytearray()
    while pos + 2 <= len(data):
        header = int.from_bytes(data[pos : pos + 2], "little")
        pos += 2
        chunk_size = (header & 0x0FFF) + 3
        is_compressed = bool(header & 0x8000)
        chunk_end = min(pos + chunk_size - 2, len(data))
        chunk_data = data[pos:chunk_end]
        pos = chunk_end
        if not is_compressed:
            decompressed.extend(chunk_data)
            continue
        cpos = 0
        while cpos < len(chunk_data):
            flag_byte = chunk_data[cpos]
            cpos += 1
            for bit in range(8):
                if cpos >= len(chunk_data):
                    break
                if (flag_byte >> bit) & 1 == 0:
                    decompressed.append(chunk_data[cpos])
                    cpos += 1
                else:
                    if cpos + 2 > len(chunk_data):
                        break
                    token = int.from_bytes(chunk_data[cpos : cpos + 2], "little")
                    cpos += 2
                    d_len = len(decompressed)
                    bits = 4
                    while (1 << bits) < d_len and bits < 12:
                        bits += 1
                    length = (token & ((1 << (16 - bits)) - 1)) + 3
                    offset = (token >> (16 - bits)) + 1
                    start = len(decompressed) - offset
                    if start >= 0:
                        for i in range(length):
                            if start + i < len(decompressed):
                                decompressed.append(decompressed[start + i])
    return bytes(decompressed)


def extract_macros_from_vba_bin(
    vba_bin: bytes, file_path: str, source_type: str = "template"
) -> List[Dict[str, Any]]:
    """Extract macro definitions from a vbaProject.bin binary blob.

    Inspects module streams under VBA/ and project symbol tables (_VBA_PROJECT).

    Args:
        vba_bin: Binary content of word/vbaProject.bin.
        file_path: Original path to the container file (.dotm / .docm).
        source_type: Label for origin ('document', 'template', 'addin_template', etc.).

    Returns:
        List of macro dictionaries with name, module, qualified_name, source, source_type, path.
    """
    if olefile is None:
        return []

    macros: List[Dict[str, Any]] = []
    seen: set = set()
    source_basename = os.path.basename(file_path)

    try:
        ole = olefile.OleFileIO(io.BytesIO(vba_bin))
        streams = ole.listdir()

        # Pass 1: Parse module streams (e.g. ['VBA', 'Module1'], ['VBA', 'EndNoteMacros'])
        for entry in streams:
            if len(entry) >= 2 and entry[0] == "VBA" and entry[-1] not in ("_VBA_PROJECT", "dir", "ThisDocument"):
                mod_name = entry[-1]
                try:
                    raw_data = ole.openstream(entry).read()
                    
                    # Search uncompressed/ASCII text
                    matches = list(re.finditer(
                        rb"(?:Public\s+|Friend\s+)?(?:Sub|Function)\s+([A-Za-z0-9_]{3,})",
                        raw_data,
                        re.IGNORECASE,
                    ))
                    
                    # If few matches, try MS-OVBA decompression
                    if not matches:
                        dec_data = decompress_ovba_chunk(raw_data)
                        if dec_data:
                            matches = list(re.finditer(
                                rb"(?:Public\s+|Friend\s+)?(?:Sub|Function)\s+([A-Za-z0-9_]{3,})",
                                dec_data,
                                re.IGNORECASE,
                            ))

                    for m in matches:
                        sub_name = m.group(1).decode("latin1", errors="ignore")
                        if not sub_name.startswith("_") and (mod_name, sub_name) not in seen:
                            seen.add((mod_name, sub_name))
                            macros.append({
                                "name": sub_name,
                                "module": mod_name,
                                "qualified_name": f"{mod_name}.{sub_name}",
                                "source": source_basename,
                                "source_type": source_type,
                                "path": file_path,
                            })
                except Exception:
                    continue

        # Pass 2: Project symbol table fallback for compiled/protected add-in modules (e.g., Zotero)
        for entry in streams:
            if entry[-1] in ("_VBA_PROJECT", "dir"):
                try:
                    stream_data = ole.openstream(entry).read()
                    words = re.findall(rb"[A-Za-z][A-Za-z0-9_]{3,}", stream_data)
                    for w in words:
                        name_str = w.decode("latin1", errors="ignore")
                        # Match common add-in macro patterns if not already found
                        if (name_str.startswith("Zotero") or name_str.startswith("EN")) and (
                            not name_str.endswith(
                                ("Ribbon", "Label", "Window", "Document", "Version", "Agent", "Command", "MessageWindow", "Tag")
                            )
                        ):
                            mod_name = "Zotero" if name_str.startswith("Zotero") else "EndNoteMacros"
                            if (mod_name, name_str) not in seen and not any(m["name"] == name_str for m in macros):
                                seen.add((mod_name, name_str))
                                macros.append({
                                    "name": name_str,
                                    "module": mod_name,
                                    "qualified_name": f"{mod_name}.{name_str}",
                                    "source": source_basename,
                                    "source_type": source_type,
                                    "path": file_path,
                                })
                except Exception:
                    continue

    except Exception:
        pass

    return macros


def extract_macros_from_file(file_path: str, source_type: str = "template") -> List[Dict[str, Any]]:
    """Extract all macros from an OpenXML .dotm or .docm file on disk.

    Args:
        file_path: Absolute path to the .dotm or .docm file.
        source_type: Label describing the role of this file.

    Returns:
        List of macro dictionaries.
    """
    if not os.path.exists(file_path) or os.path.basename(file_path).startswith("~$"):
        return []

    try:
        with zipfile.ZipFile(file_path, "r") as z:
            if "word/vbaProject.bin" not in z.namelist():
                return []
            vba_bin = z.read("word/vbaProject.bin")
            return extract_macros_from_vba_bin(vba_bin, file_path, source_type=source_type)
    except Exception:
        return []


def discover_template_files(app=None, doc=None) -> List[Tuple[str, str]]:
    """Discover all active document, template, and add-in files that may contain Word macros.

    Args:
        app: Word.Application COM object (optional).
        doc: Document COM object (optional).

    Returns:
        List of (file_path, source_type) tuples.
    """
    candidates: List[Tuple[str, str]] = []
    seen_paths: set = set()

    def add_candidate(path: str, src_type: str):
        if not path:
            return
        norm = os.path.normcase(os.path.abspath(path))
        if norm not in seen_paths and os.path.exists(path) and not os.path.basename(path).startswith("~$"):
            seen_paths.add(norm)
            candidates.append((os.path.abspath(path), src_type))

    # 1. Active Document (if macro-enabled)
    if doc is not None:
        try:
            doc_path = doc.FullName
            if doc_path and doc_path.lower().endswith((".docm", ".dotm")):
                add_candidate(doc_path, "document")
        except Exception:
            pass

        # 2. Attached Template
        try:
            attached = doc.AttachedTemplate
            if attached is not None:
                att_path = attached.FullName if hasattr(attached, "FullName") else os.path.join(attached.Path, attached.Name)
                add_candidate(att_path, "attached_template")
        except Exception:
            pass

    # 3. Normal.dotm
    normal_path = os.path.expandvars(r"%APPDATA%\Microsoft\Templates\Normal.dotm")
    add_candidate(normal_path, "normal_template")

    # 4. Word STARTUP Folder
    startup_dir = os.path.expandvars(r"%APPDATA%\Microsoft\Word\STARTUP")
    if os.path.exists(startup_dir):
        for ext in ("*.dotm", "*.docm", "*.dot"):
            for p in glob.glob(os.path.join(startup_dir, ext)):
                add_candidate(p, "startup_template")

    # 5. Loaded AddIns in Word.Application
    if app is not None:
        try:
            for i in range(1, app.AddIns.Count + 1):
                addin = app.AddIns(i)
                p = os.path.join(addin.Path, addin.Name)
                add_candidate(p, "addin_template")
        except Exception:
            pass

    # 6. Common Add-In Companion Directories (e.g. EndNote CWYW)
    common_patterns = [
        r"C:\Program Files\Common Files\ResearchSoft\Cwyw\*\EndNote Cwyw.dotm",
        r"C:\Program Files (x86)\Common Files\ResearchSoft\Cwyw\*\EndNote Cwyw.dotm",
        r"C:\Program Files\Common Files\ResearchSoft\Cwyw\*\EndNote Online Cwyw.dotm",
        r"C:\Program Files (x86)\Common Files\ResearchSoft\Cwyw\*\EndNote Online Cwyw.dotm",
    ]
    for pattern in common_patterns:
        for p in glob.glob(pattern):
            add_candidate(p, "addin_template")

    return candidates


def list_available_macros(
    app=None,
    doc=None,
    query: Optional[str] = None,
    source_filter: str = "all",
) -> List[Dict[str, Any]]:
    """Enumerate all available macros across document, templates, and add-ins.

    First attempts Word VBE COM reflection if trusted, then falls back to file parsing.

    Args:
        app: Word.Application COM object.
        doc: Document COM object.
        query: Optional case-insensitive filter for macro name, module, or source.
        source_filter: 'all' | 'document' | 'templates' | 'addins'

    Returns:
        Deduplicated, sorted list of macro entries.
    """
    all_macros: List[Dict[str, Any]] = []
    seen: set = set()

    # Pass 1: Try COM VBE inspection if programmatic access is enabled
    vbe_worked = False
    if app is not None:
        try:
            vbe = app.VBE
            if vbe and vbe.VBProjects:
                for proj in vbe.VBProjects:
                    proj_name = getattr(proj, "Name", "VBAProject")
                    for comp in proj.VBComponents:
                        code_mod = getattr(comp, "CodeModule", None)
                        if code_mod and hasattr(code_mod, "CountOfLines"):
                            line_count = code_mod.CountOfLines
                            mod_name = getattr(comp, "Name", "Module")
                            line_idx = 1
                            while line_idx <= line_count:
                                # ProcOfLine(line, kind=0 for standard sub/proc)
                                try:
                                    proc_name = code_mod.ProcOfLine(line_idx, 0)
                                    if proc_name and (mod_name, proc_name) not in seen:
                                        seen.add((mod_name, proc_name))
                                        all_macros.append({
                                            "name": proc_name,
                                            "module": mod_name,
                                            "qualified_name": f"{mod_name}.{proc_name}",
                                            "source": proj_name,
                                            "source_type": "vbe",
                                            "path": "",
                                        })
                                        vbe_worked = True
                                        # Skip lines belonging to this procedure
                                        line_idx += max(1, code_mod.ProcCountLines(proc_name, 0))
                                        continue
                                except Exception:
                                    pass
                                line_idx += 1
        except Exception:
            vbe_worked = False

    # Pass 2: File-based extraction from templates and add-ins
    files = discover_template_files(app, doc)
    for file_path, src_type in files:
        file_macros = extract_macros_from_file(file_path, source_type=src_type)
        for m in file_macros:
            key = (m["module"], m["name"])
            if key not in seen:
                seen.add(key)
                all_macros.append(m)

    # Filtering by source
    filtered: List[Dict[str, Any]] = []
    src_filter_clean = source_filter.strip().lower() if source_filter else "all"

    for m in all_macros:
        st = m.get("source_type", "").lower()
        if src_filter_clean == "document" and st != "document":
            continue
        elif src_filter_clean == "templates" and "template" not in st:
            continue
        elif src_filter_clean == "addins" and "addin" not in st and "startup" not in st:
            continue

        # Filter by search query if provided
        if query:
            q = query.strip().lower()
            name_match = q in m["name"].lower()
            mod_match = q in m["module"].lower()
            src_match = q in m["source"].lower()
            qual_match = q in m["qualified_name"].lower()
            if not (name_match or mod_match or src_match or qual_match):
                continue

        filtered.append(m)

    # Sort deterministically by source and name
    filtered.sort(key=lambda x: (x.get("source", ""), x.get("name", "")))
    return filtered


def ensure_addin_template_connected(app, macro_name: str) -> Dict[str, Any]:
    """Ensure that the add-in template providing macro_name is loaded and enabled in Word.

    Specifically handles EndNote CWYW (EndNote Cwyw.dotm) and Zotero (Zotero.dotm).

    Args:
        app: Word.Application COM object.
        macro_name: Name of the macro to be executed.

    Returns:
        Dict with status and details on any add-in loaded or enabled.
    """
    if app is None or not macro_name:
        return {"connected": False}

    name_lower = macro_name.lower()
    is_endnote = name_lower.startswith("en") or "endnote" in name_lower
    is_zotero = name_lower.startswith("zotero")

    if not (is_endnote or is_zotero):
        return {"connected": False}

    target_name = "endnote" if is_endnote else "zotero"

    try:
        # Check if already registered in AddIns
        for i in range(1, app.AddIns.Count + 1):
            addin = app.AddIns(i)
            if target_name in addin.Name.lower():
                if not addin.Installed:
                    addin.Installed = True
                    return {"connected": True, "action": "enabled_existing", "name": addin.Name}
                return {"connected": True, "action": "already_installed", "name": addin.Name}

        # If not registered in AddIns collection, search common install locations
        if is_endnote:
            patterns = [
                r"C:\Program Files\Common Files\ResearchSoft\Cwyw\*\EndNote Cwyw.dotm",
                r"C:\Program Files (x86)\Common Files\ResearchSoft\Cwyw\*\EndNote Cwyw.dotm",
            ]
            for pattern in patterns:
                for p in glob.glob(pattern):
                    if os.path.exists(p):
                        new_addin = app.AddIns.Add(FileName=p, Install=True)
                        return {"connected": True, "action": "added_and_installed", "path": p, "name": new_addin.Name}

        if is_zotero:
            startup_path = os.path.expandvars(r"%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm")
            if os.path.exists(startup_path):
                new_addin = app.AddIns.Add(FileName=startup_path, Install=True)
                return {"connected": True, "action": "added_and_installed", "path": startup_path, "name": new_addin.Name}

    except Exception as e:
        return {"connected": False, "error": str(e)}

    return {"connected": False}
