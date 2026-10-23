import asyncio
import json
import pytest
from word_document_server.tools.live_read_tools import (
    word_live_list_fields,
    word_live_get_field,
)
from word_document_server.tools.live_tools import (
    word_live_update_field_code,
)


@pytest.mark.anyio
async def test_mcptest_fields():
    # 0. Normalization check: Ensure Field 1 is restored to 105 if previous run crashed
    init_res_raw = await word_live_get_field(filename="mcptest.docx", field_index=1)
    init_res = json.loads(init_res_raw)
    if init_res.get("success") and "<RecNum>999</RecNum>" in init_res.get("field", {}).get("code", ""):
        restored = init_res["field"]["code"].replace("<RecNum>999</RecNum>", "<RecNum>105</RecNum>").replace("<rec-number>999</rec-number>", "<rec-number>105</rec-number>")
        await word_live_update_field_code(filename="mcptest.docx", field_index=1, new_code=restored)

    # 1. List all fields in mcptest.docx
    res_raw = await word_live_list_fields(filename="mcptest.docx")
    res = json.loads(res_raw)
    print("\n--- All Fields ---")
    print(json.dumps(res, indent=2))
    assert res.get("success") is True, f"Error: {res}"
    assert res.get("matched_fields", 0) > 0, "Expected at least 1 field in mcptest.docx"

    # 2. Filter by cite
    res_cite_raw = await word_live_list_fields(filename="mcptest.docx", field_type="cite")
    res_cite = json.loads(res_cite_raw)
    print("\n--- Cite Fields ---")
    print(json.dumps(res_cite, indent=2))
    assert res_cite.get("success") is True
    assert res_cite.get("matched_fields", 0) >= 3, "Expected at least 3 citations [1], [2], [3]"

    # Verify parsed EndNote metadata
    first_cite = res_cite["fields"][0]
    assert "endnote" in first_cite, "First cite should have parsed EndNote data"
    assert len(first_cite["endnote"]["record_numbers"]) > 0, "Should have at least 1 record number"
    print(f"First cite record number: {first_cite['endnote']['record_numbers']}")

    # 3. Test get_field
    get_res_raw = await word_live_get_field(filename="mcptest.docx", field_index=1)
    get_res = json.loads(get_res_raw)
    print("\n--- Field #1 Details ---")
    print(json.dumps(get_res, indent=2))
    assert get_res.get("success") is True
    assert get_res["field"]["index"] == 1
    assert "code" in get_res["field"]
    assert len(get_res["field"]["code"]) > 0

    # 4. Filter by contains (record number or author in code)
    rec_num = str(first_cite["endnote"]["record_numbers"][0])
    res_contains_raw = await word_live_list_fields(filename="mcptest.docx", contains=rec_num)
    res_contains = json.loads(res_contains_raw)
    assert res_contains.get("success") is True
    assert res_contains.get("matched_fields", 0) >= 1

    # 5. Test word_live_update_field_code and revert
    original_code = get_res["field"]["code"]
    modified_code = (
        original_code
        .replace("<RecNum>105</RecNum>", "<RecNum>999</RecNum>")
        .replace("<rec-number>105</rec-number>", "<rec-number>999</rec-number>")
    )
    try:
        update_res_raw = await word_live_update_field_code(filename="mcptest.docx", field_index=1, new_code=modified_code)
        update_res = json.loads(update_res_raw)
        print("\n--- Update Field #1 ---")
        print(json.dumps(update_res, indent=2))
        assert update_res.get("success") is True
        assert 999 in update_res["endnote"]["record_numbers"]

        # Verify via get_field that change took effect in live Word memory
        verify_res_raw = await word_live_get_field(filename="mcptest.docx", field_index=1)
        verify_res = json.loads(verify_res_raw)
        print("\n--- Verify Field #1 ---")
        print(json.dumps(verify_res, indent=2))
        assert 999 in verify_res.get("field", {}).get("endnote", {}).get("record_numbers", [])
    finally:
        # Revert back to original code to preserve document state
        revert_res_raw = await word_live_update_field_code(filename="mcptest.docx", field_index=1, new_code=original_code)
        revert_res = json.loads(revert_res_raw)
        assert revert_res.get("success") is True
        assert 105 in revert_res["endnote"]["record_numbers"]
        print("\nSuccessfully reverted Field #1 back to original code.")


if __name__ == "__main__":
    asyncio.run(test_mcptest_fields())
