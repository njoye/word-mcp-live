import os
import pytest
from word_document_server.utils.macro_utils import (
    discover_template_files,
    list_available_macros,
    ensure_addin_template_connected,
    decompress_ovba_chunk,
)


def test_discover_template_files():
    templates = discover_template_files(app=None, doc=None)
    assert isinstance(templates, list)
    # On this machine, EndNote and/or Zotero templates exist
    assert len(templates) > 0
    for path, src_type in templates:
        assert os.path.exists(path)
        assert src_type in ("document", "attached_template", "normal_template", "startup_template", "addin_template")


def test_list_available_macros_all():
    macros = list_available_macros(app=None, doc=None)
    assert isinstance(macros, list)
    assert len(macros) > 0
    first = macros[0]
    assert "name" in first
    assert "module" in first
    assert "source" in first
    assert "source_type" in first


def test_list_available_macros_query_filter():
    biblio_macros = list_available_macros(app=None, doc=None, query="biblio")
    assert len(biblio_macros) > 0
    names = [m["name"] for m in biblio_macros]
    assert any("Bibliography" in name or "FormatBibliography" in name for name in names)


def test_list_available_macros_source_filter():
    addin_macros = list_available_macros(app=None, doc=None, source_filter="addins")
    assert len(addin_macros) > 0
    for m in addin_macros:
        assert m["source_type"] in ("addin_template", "startup_template")


def test_ensure_addin_template_connected_none_app():
    # If app is None, safely returns connected=False
    res = ensure_addin_template_connected(None, "ENFormatBibliography")
    assert res == {"connected": False}

    res_unknown = ensure_addin_template_connected(None, "SomeRandomCustomMacro")
    assert res_unknown == {"connected": False}


def test_decompress_ovba_chunk_empty():
    assert decompress_ovba_chunk(b"") == b""
