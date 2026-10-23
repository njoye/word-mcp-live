"""Utilities for inspecting, parsing, and classifying Word fields and EndNote citation codes."""

import html
import re

# Word COM WdFieldType enumerations
FIELD_TYPES = {
    -1: "Empty",
    3: "Ref",
    4: "IndexEntry",
    5: "FootnoteRef",
    6: "Set",
    7: "If",
    8: "Index",
    9: "TOCEntry",
    10: "StyleRef",
    11: "RefDoc",
    12: "Sequence",
    13: "TOC",
    14: "Info",
    15: "Title",
    16: "Subject",
    17: "Author",
    18: "KeyWord",
    19: "Comments",
    20: "LastSavedBy",
    21: "CreateDate",
    22: "SaveDate",
    23: "PrintDate",
    24: "RevisionNum",
    25: "EditTime",
    26: "NumPages",
    27: "NumWords",
    28: "NumChars",
    29: "FileName",
    30: "Template",
    31: "Date",
    32: "Time",
    33: "Page",
    34: "Expression",
    35: "Quote",
    36: "Include",
    37: "PageRef",
    38: "Ask",
    39: "FillIn",
    40: "Data",
    41: "Next",
    42: "NextIf",
    43: "SkipIf",
    44: "MergeRec",
    45: "DDE",
    46: "DDEAuto",
    47: "Glossary",
    48: "Print",
    49: "Formula",
    50: "GoToButton",
    51: "MacroButton",
    52: "AutoNumOutline",
    53: "AutoNumLegal",
    54: "AutoNum",
    55: "Import",
    56: "Link",
    57: "Symbol",
    58: "Embed",
    59: "MergeField",
    60: "UserName",
    61: "UserInitials",
    62: "UserAddress",
    63: "BarCode",
    64: "DocVariable",
    65: "Section",
    66: "SectionPages",
    67: "IncludePicture",
    68: "IncludeText",
    69: "FileSize",
    70: "FormTextInput",
    71: "FormCheckBox",
    72: "NoteRef",
    73: "TOA",
    74: "TOAEntry",
    75: "MergeSeq",
    77: "Private",
    78: "Database",
    79: "AutoText",
    80: "Compare",
    81: "Addin",
    82: "Subscriber",
    83: "FormDropDown",
    84: "Advance",
    85: "DocProperty",
    87: "OCX",
    88: "Hyperlink",
    89: "AutoTextList",
    90: "ListNum",
    91: "HTMLActiveX",
    92: "BidiOutline",
    93: "AddressBlock",
    94: "GreetingLine",
    95: "Shape",
    96: "Citation",
    97: "Bibliography",
    98: "MergeBarcode",
    99: "DisplayBarcode",
    100: "Copilot",
}


def get_field_type_name(type_id: int) -> str:
    """Return friendly name for a WdFieldType integer."""
    return FIELD_TYPES.get(type_id, f"FieldType_{type_id}")


def parse_endnote_field_code(code_text: str) -> dict | None:
    """Parse EndNote citation metadata from a field code string.

    Supports:
    - Modern EndNote CWYW XML (ADDIN EN.CITE <EndNote><Cite>...)
    - EndNote CWYW unformatted/temporary citations ({Author, Year #RecNum})
    - Legacy/alternate Cite formats (Cite "library.enl" at="RecordIndex=...")
    - EndNote bibliography fields (ADDIN EN.REFLIST)

    Returns a dict with parsed metadata or None if not an EndNote field.
    """
    if not code_text:
        return None

    code_clean = code_text.strip()
    is_reflist = bool(re.search(r'\bEN\.REFLIST\b', code_clean, re.I))
    is_cite = bool(re.search(r'\bEN\.CITE\b|\bEN\.CITE\.DATA\b|\bCite\b.*RecordIndex', code_clean, re.I))
    is_temp_cite = bool(re.search(r'\{[^{}]*#\d+[^{}]*\}', code_clean))

    if not (is_cite or is_reflist or is_temp_cite):
        return None

    if is_reflist:
        return {
            "is_endnote": True,
            "field_category": "bibliography",
            "record_numbers": [],
            "authors": [],
            "years": [],
            "titles": [],
            "library": None,
        }

    # Extract Record Numbers
    rec_nums = []
    # 1. XML tags: <RecNum>168</RecNum> or <rec-number>168</rec-number>
    for match in re.findall(r'<(?:RecNum|rec-number)>(\d+)</(?:RecNum|rec-number)>', code_clean, re.I):
        try:
            num = int(match)
            if num not in rec_nums:
                rec_nums.append(num)
        except ValueError:
            pass

    # 2. Key-value: RecordIndex=168 or RecordIndex="168"
    for match in re.findall(r'RecordIndex=["\']?(\d+)["\']?', code_clean, re.I):
        try:
            num = int(match)
            if num not in rec_nums:
                rec_nums.append(num)
        except ValueError:
            pass

    # 3. Temp citation: #168
    if not rec_nums:
        for match in re.findall(r'#(\d+)', code_clean):
            try:
                num = int(match)
                if num not in rec_nums:
                    rec_nums.append(num)
            except ValueError:
                pass

    # Extract Authors
    authors = []
    for match in re.findall(r'<(?:Author|author)>(.*?)</(?:Author|author)>', code_clean, re.I | re.DOTALL):
        cleaned_author = re.sub(r'<[^>]+>', '', match).strip()
        cleaned_author = html.unescape(cleaned_author)
        if cleaned_author and cleaned_author not in authors:
            authors.append(cleaned_author)

    if not authors:
        for match in re.findall(r'(?:Author|author_hint)=["\']([^"\']+)["\']', code_clean, re.I):
            cleaned_author = html.unescape(match.strip())
            if cleaned_author and cleaned_author not in authors:
                authors.append(cleaned_author)

    # Extract Years
    years = []
    for match in re.findall(r'<(?:Year|year)>(.*?)</(?:Year|year)>', code_clean, re.I | re.DOTALL):
        cleaned_year = re.sub(r'<[^>]+>', '', match).strip()
        cleaned_year = html.unescape(cleaned_year)
        if cleaned_year and cleaned_year not in years:
            years.append(cleaned_year)

    if not years:
        for match in re.findall(r'(?:Year|year_hint)=["\']([^"\']+)["\']', code_clean, re.I):
            cleaned_year = html.unescape(match.strip())
            if cleaned_year and cleaned_year not in years:
                years.append(cleaned_year)

    # Extract Titles
    titles = []
    for match in re.findall(r'<(?:Title|title)>(.*?)</(?:Title|title)>', code_clean, re.I | re.DOTALL):
        cleaned_title = re.sub(r'<[^>]+>', '', match).strip()
        cleaned_title = html.unescape(cleaned_title)
        if cleaned_title and cleaned_title not in titles:
            titles.append(cleaned_title)

    # Extract Library name if present
    library = None
    lib_match = re.search(r'(?:Cite\s+|library=)["\']([^"\']+\.enl)["\']', code_clean, re.I)
    if lib_match:
        library = lib_match.group(1)
    else:
        db_match = re.search(r'<database>([^<]+\.enl)</database>', code_clean, re.I)
        if db_match:
            library = db_match.group(1)

    return {
        "is_endnote": True,
        "field_category": "citation",
        "record_numbers": rec_nums,
        "authors": authors,
        "years": years,
        "titles": titles,
        "library": library,
    }
