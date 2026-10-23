import json
import pytest
from word_document_server.tools.live_read_tools import word_live_list_macros
from word_document_server.tools.live_tools import word_live_run_macro


@pytest.mark.anyio
async def test_live_list_and_run_macros():
    # 1. Test listing all macros
    res_raw = await word_live_list_macros(filename="mcptest.docx")
    data = json.loads(res_raw)
    assert data.get("success") is True
    assert data.get("total_macros", 0) > 0
    assert "macros" in data
    assert len(data["macros"]) > 0

    # 2. Test filtering by query
    res_filter_raw = await word_live_list_macros(filename="mcptest.docx", filter="biblio")
    filter_data = json.loads(res_filter_raw)
    assert filter_data.get("success") is True
    macro_names = [m["name"] for m in filter_data["macros"]]
    # Should include ENFormatBibliography or ZoteroAddEditBibliography
    assert any("FormatBibliography" in name or "Bibliography" in name for name in macro_names)

    # 3. Test empty macro validation
    empty_raw = await word_live_run_macro(macro_name="")
    empty_data = json.loads(empty_raw)
    assert "error" in empty_data
    assert "macro_name must be specified" in empty_data["error"]

    # 4. Test macro error handling and hint
    invalid_raw = await word_live_run_macro(macro_name="NonExistentMacro_12345", filename="mcptest.docx")
    invalid_data = json.loads(invalid_raw)
    assert "error" in invalid_data
    if invalid_data.get("hint"):
        assert "word_live_list_macros" in invalid_data["hint"]
