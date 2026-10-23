import pytest
from word_document_server.utils.field_utils import (
    get_field_type_name,
    parse_endnote_field_code,
)


def test_get_field_type_name():
    assert get_field_type_name(81) == "Addin"
    assert get_field_type_name(26) == "NumPages"
    assert get_field_type_name(33) == "Page"
    assert get_field_type_name(13) == "TOC"
    assert get_field_type_name(88) == "Hyperlink"
    assert get_field_type_name(96) == "Citation"
    assert get_field_type_name(97) == "Bibliography"
    assert get_field_type_name(9999) == "FieldType_9999"


def test_parse_endnote_xml_citation():
    code = (
        'ADDIN EN.CITE <EndNote><Cite><Author>Jocher</Author><Year>2026</Year>'
        '<RecNum>168</RecNum><DisplayText>[1]</DisplayText>'
        '<record><title>Ultralytics YOLO26</title></record></Cite></EndNote>'
    )
    parsed = parse_endnote_field_code(code)
    assert parsed is not None
    assert parsed["is_endnote"] is True
    assert parsed["field_category"] == "citation"
    assert parsed["record_numbers"] == [168]
    assert parsed["authors"] == ["Jocher"]
    assert parsed["years"] == ["2026"]
    assert parsed["titles"] == ["Ultralytics YOLO26"]


def test_parse_endnote_multi_citation():
    code = (
        'ADDIN EN.CITE <EndNote>'
        '<Cite><Author>Smith</Author><Year>2024</Year><RecNum>42</RecNum></Cite>'
        '<Cite><Author>Jones &amp; Brown</Author><Year>2025</Year><RecNum>99</RecNum></Cite>'
        '</EndNote>'
    )
    parsed = parse_endnote_field_code(code)
    assert parsed is not None
    assert parsed["record_numbers"] == [42, 99]
    assert parsed["authors"] == ["Smith", "Jones & Brown"]
    assert parsed["years"] == ["2024", "2025"]


def test_parse_endnote_legacy_format():
    code = 'Cite "v2x.enl" at="RecordIndex=168 Guide=\'EndNote 2026\' Author=\'Jocher\' Year=\'2026\'"'
    parsed = parse_endnote_field_code(code)
    assert parsed is not None
    assert parsed["record_numbers"] == [168]
    assert parsed["library"] == "v2x.enl"
    assert parsed["authors"] == ["Jocher"]
    assert parsed["years"] == ["2026"]


def test_parse_endnote_temp_citation():
    code = '{Nateghinia, 2026 #1}'
    parsed = parse_endnote_field_code(code)
    assert parsed is not None
    assert parsed["record_numbers"] == [1]


def test_parse_endnote_reflist():
    code = 'ADDIN EN.REFLIST'
    parsed = parse_endnote_field_code(code)
    assert parsed is not None
    assert parsed["is_endnote"] is True
    assert parsed["field_category"] == "bibliography"


def test_parse_non_endnote_field():
    assert parse_endnote_field_code('PAGE') is None
    assert parse_endnote_field_code('TOC \\o "1-3"') is None
    assert parse_endnote_field_code('') is None
    assert parse_endnote_field_code(None) is None
