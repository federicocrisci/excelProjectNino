#!/usr/bin/env python3
"""Build the offline converter model from the user's paired workbooks.

Development utility only. End users need no Python or LibreOffice. Supply a trusted
LibreOffice conversion of the XLS reference through --converted; styles are taken
from its raw OOXML, while reference cell values are read from the original XLS.
Original input documents need not be copied into the project or its distributable.
"""
from __future__ import annotations

import argparse
import base64
from collections import Counter, defaultdict
from difflib import SequenceMatcher
import io
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

import openpyxl
import xlrd

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"s": MAIN}


def numeric(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def code(value):
    if numeric(value) and float(value).is_integer():
        return str(int(value))
    return str(value).strip() if value is not None else ""


def source_rows(path):
    book = openpyxl.load_workbook(path, data_only=False)
    if len(book.worksheets) != 1:
        raise ValueError("The model builder expects the selected one-sheet source.")
    sheet = book.active
    records = []
    all_rows = []
    for row in sheet.iter_rows():
        values = [cell.value for cell in row[:18]]
        values.extend([None] * (18 - len(values)))
        all_rows.append((row[0].row, values))
        if values[1] == "I" and code(values[2]) and values[6] is not None:
            records.append((row[0].row, values))
    return sheet.title, records, all_rows


def match_products(source, target):
    """Keep duplicate occurrences and distinguish gifts from regular products."""
    by_code = defaultdict(list)
    for row_number, values in source:
        by_code[code(values[2])].append((row_number, values))
    used = set()
    matched = {}
    for row_number, values in target:
        if not code(values[2]):
            continue
        candidates = [record for record in by_code[code(values[2])] if record[0] not in used]
        if not candidates:
            raise ValueError(f"Reference product row {row_number} has no source occurrence")

        def score(record):
            _, original = record
            description = " ".join(str(v).strip() for v in original[6:11] if v is not None)
            final_description = str(values[5])
            same_gift = ("OMAGGIO" in description) == ("OMAGGIO" in final_description)
            return SequenceMatcher(None, description, final_description).ratio() + int(same_gift)

        selected = max(candidates, key=score)
        used.add(selected[0])
        matched[row_number] = selected
    excluded = []
    for row_number, values in source:
        if row_number not in used:
            reason = (
                "Omaggio aggiuntivo accorpato alla riga del kit precedente nel modello di riferimento."
                if str(values[6]).lstrip().startswith("+")
                else "Omaggio non presente nel file lavorato di riferimento."
            )
            excluded.append({"sourceRow": row_number, "source": values, "reason": reason})
    return matched, excluded


def collect_source_details(all_source, products):
    """Attach buying conditions and display composition to their real products."""
    details = []
    product_numbers = {row for row, _ in products}
    for source_row, values in all_source:
        if source_row in product_numbers or code(values[2]) or not values[0]:
            continue
        if not any(value is not None and value != "" for value in values[6:11]):
            continue
        condition = isinstance(values[6], str) and values[6].lstrip().upper().startswith("ACQUISTANDO")
        siblings = [
            row for row, product in products
            if product[0] == values[0]
            and (row > source_row if condition else row < source_row)
        ]
        parent = (min(siblings) if condition else max(siblings)) if siblings else None
        details.append({
            "sourceRow": source_row, "source": values, "parentSourceRow": parent,
            "kind": "condition" if condition else "composition",
        })
    return details


def attach_reference_details(rows, source_details):
    """Associate each reference detail with its product and original detail row."""
    used = set()
    for index, row in enumerate(rows):
        if row["kind"] != "detail":
            continue
        condition = str(row["values"][5]).lstrip().upper().startswith("ACQUISTANDO")
        candidates = rows[index + 1:] if condition else reversed(rows[:index])
        parent = next((candidate for candidate in candidates if candidate["kind"] == "product"), None)
        if parent is None:
            raise ValueError(f"Reference detail row {row['r']} has no parent product")
        row["detailKind"] = "condition" if condition else "composition"
        row["parentR"] = parent["r"]
        row["detailParentR"] = parent["r"]
        details = [
            item for item in source_details
            if item["sourceRow"] not in used
            and item["parentSourceRow"] == parent["sourceRow"]
            and item["kind"] == row["detailKind"]
        ]
        if not details:
            raise ValueError(f"Reference detail row {row['r']} has no source detail")

        def score(item):
            description = " ".join(str(v).strip() for v in item["source"][6:11] if v is not None)
            return SequenceMatcher(None, description, str(row["values"][5])).ratio()

        selected = max(details, key=score)
        used.add(selected["sourceRow"])
        row["sourceDetailRow"] = selected["sourceRow"]


def strip_template(converted, original_sheet):
    """Preserve styles/widths/printing, remove data and stale external dependencies."""
    with zipfile.ZipFile(converted) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    sheet_path = "xl/worksheets/sheet1.xml"
    root = ET.fromstring(files[sheet_path])
    style_rows = {}
    for row in root.findall("s:sheetData/s:row", NS):
        row_number = int(row.attrib["r"])
        styles = [0] * 13
        for cell in row.findall("s:c", NS):
            letters = re.match(r"[A-Z]+", cell.attrib["r"]).group()
            column = 0
            for letter in letters:
                column = column * 26 + ord(letter) - ord("A") + 1
            if column <= 13:
                styles[column - 1] = int(cell.attrib.get("s", "0"))
        style_rows[row_number] = {"attrs": dict(row.attrib), "styles": styles}
    sheet = files[sheet_path].decode("utf-8")
    sheet, changed = re.subn(r"<sheetData>.*?</sheetData>", "<sheetData></sheetData>", sheet, flags=re.S)
    if changed != 1:
        raise ValueError("Could not replace exactly one reference sheetData")
    sheet = re.sub(r'<dimension\s+ref="[^"]+"\s*/>', '<dimension ref="A1:M6619"/>', sheet)
    # LibreOffice rounds exported widths to two decimals. Recover BIFF's exact
    # units of 1/256 character so the new workbook retains the reference layout.
    def precise_width(match):
        element = match.group()
        first = int(re.search(r'\bmin="(\d+)"', element).group(1))
        last = int(re.search(r'\bmax="(\d+)"', element).group(1))
        if first != last or not 1 <= first <= 13:
            return element
        original = original_sheet.colinfo_map.get(first - 1)
        if original is None:
            raise ValueError(f"Reference column {first} has no width information")
        width = format(original.width / 256, ".15g")
        return re.sub(r'\bwidth="[^"]+"', f'width="{width}"', element)

    sheet = re.sub(r"<col\b[^>]*/>", precise_width, sheet)
    # Invalid saved filters/conditional ranges should not refer to old row counts.
    sheet = re.sub(r"<autoFilter\b[^>]*/>|<autoFilter\b[^>]*>.*?</autoFilter>", "", sheet, flags=re.S)
    files[sheet_path] = sheet.encode("utf-8")
    workbook_path = "xl/workbook.xml"
    workbook = files[workbook_path].decode("utf-8")
    workbook = re.sub(r"<externalReferences\b[^>]*>.*?</externalReferences>", "", workbook, flags=re.S)

    def print_names(match):
        entries = re.findall(r"<definedName\b[^>]*>.*?</definedName>", match.group(), flags=re.S)
        entries = [entry for entry in entries if re.search(r'name="_xlnm\.Print_(?:Area|Titles)"', entry)]
        return "<definedNames>" + "".join(entries) + "</definedNames>" if entries else ""

    workbook = re.sub(r"<definedNames\b[^>]*>.*?</definedNames>", print_names, workbook, flags=re.S)
    workbook = re.sub(r"<calcPr\b[^>]*/>", '<calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/>', workbook)
    files[workbook_path] = workbook.encode("utf-8")
    for name in list(files):
        if name.startswith("xl/externalLinks/") or name in {"xl/sharedStrings.xml", "xl/calcChain.xml"}:
            del files[name]
        elif name.endswith(".rels"):
            text = files[name].decode("utf-8")
            text = re.sub(
                r'<Relationship\b(?=[^>]*Type="[^"]*/(?:externalLink|sharedStrings|calcChain)")[^>]*/>',
                "", text,
            )
            files[name] = text.encode("utf-8")
    content_types = files["[Content_Types].xml"].decode("utf-8")
    content_types = re.sub(
        r'<Override\b(?=[^>]*PartName="/(?:xl/externalLinks/[^\"]+|xl/sharedStrings\.xml|xl/calcChain\.xml)")[^>]*/>',
        "", content_types,
    )
    files["[Content_Types].xml"] = content_types.encode("utf-8")
    # Remove reference authors, timestamps and paths from the reusable application.
    files["docProps/core.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>New</dc:title>'
        '<dc:creator>Canvas</dc:creator></cp:coreProperties>'
    ).encode("utf-8")
    for name, payload in files.items():
        if name.endswith(".xml") or name.endswith(".rels"):
            ET.fromstring(payload)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for name, payload in files.items():
            archive.writestr(name, payload)
    return base64.b64encode(output.getvalue()).decode("ascii"), style_rows


def build(source_path, reference_path, converted_path):
    source_sheet, source, all_source = source_rows(source_path)
    reference = xlrd.open_workbook(reference_path, formatting_info=True)
    target_sheet = reference.sheet_by_index(0)
    target = [(row + 1, target_sheet.row_values(row)) for row in range(target_sheet.nrows)]
    matches, excluded = match_products(source, target)
    source_details = collect_source_details(all_source, source)
    template_zip, styles = strip_template(converted_path, target_sheet)
    rows = []
    category = ""
    net_lookup = {}
    gross_lookup = {}
    anomalies = []
    for row_number, values in target:
        values = (values + [""] * 13)[:13]
        if values[0]:
            category = str(values[0]).strip()
        if row_number == 1:
            kind = "header"
        elif values[0]:
            kind = "category"
        elif code(values[2]):
            kind = "product"
        elif any(value != "" for value in values):
            kind = "detail"
        else:
            kind = "blank"
        row = {"r": row_number, **styles[row_number], "values": values, "kind": kind, "category": category}
        if row_number in matches:
            source_row, original = matches[row_number]
            row["source"] = original
            row["sourceRow"] = source_row
            # Avoid embedding a known tenfold typo in the rounding lookup.
            discount = original[15] if numeric(original[15]) else 0
            anomalous_price = (
                numeric(original[14]) and numeric(values[7])
                and abs(original[14] * (1 - discount) - values[7]) > 0.011
            )
            if anomalous_price:
                anomalies.append({
                    "code": code(original[2]), "sourceRow": source_row, "targetRow": row_number,
                    "sourceNet": original[14], "discount": discount, "referenceNet": values[7],
                    "reason": "Prezzo del riferimento incoerente con il listino originale; usare il prezzo sorgente.",
                })
            if numeric(original[14]) and original[15] != "SM" and numeric(values[7]) and not anomalous_price:
                pair = (original[14], discount)
                if pair in net_lookup and net_lookup[pair] != values[7]:
                    raise ValueError(f"Conflicting reference net prices for {pair}")
                net_lookup[pair] = values[7]
            if all(numeric(value) for value in [values[7], original[16], values[8]]) and not anomalous_price:
                pair = (values[7], original[16])
                if pair in gross_lookup and gross_lookup[pair] != values[8]:
                    raise ValueError(f"Conflicting reference gross prices for {pair}")
                gross_lookup[pair] = values[8]
        rows.append(row)
    attach_reference_details(rows, source_details)
    counts = Counter(row["kind"] for row in rows)
    return {
        "version": 1,
        "sourceName": source_path.name,
        "targetName": reference_path.name,
        "sourceSheet": source_sheet,
        "sheetName": target_sheet.name,
        "templateZip": template_zip,
        "rows": rows,
        "sourceDetails": source_details,
        "excluded": excluded,
        "netLookup": [[price, discount, net] for (price, discount), net in sorted(net_lookup.items())],
        "grossLookup": [[net, vat, gross] for (net, vat), gross in sorted(gross_lookup.items())],
        "genericStyles": {
            "product": styles[4]["styles"], "category": styles[8]["styles"],
            "header": styles[1]["styles"], "detail": styles[585]["styles"],
            "blank": styles[248]["styles"],
        },
        "anomalies": anomalies,
        "stats": {
            "sourceRows": len(all_source), "sourceProducts": len(source), "targetRows": len(rows),
            "targetProducts": counts["product"], "categories": counts["category"],
            "details": counts["detail"], "excludedProducts": len(excluded),
            "stockQuantitiesInReference": sum(1 for row in rows if row["kind"] == "product" and numeric(row["values"][10]) and row["values"][10] != 0),
        },
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, type=Path, help="Original XLSX")
    parser.add_argument("--reference", required=True, type=Path, help="Worked reference XLS")
    parser.add_argument("--converted", required=True, type=Path, help="Trusted XLSX conversion of reference")
    parser.add_argument("--output", type=Path, default=Path(__file__).resolve().parents[1] / "template.js")
    args = parser.parse_args()
    model = build(args.source, args.reference, args.converted)
    encoded = json.dumps(model, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    encoded = encoded.replace("</", "<\\/").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
    args.output.write_text("/* Generated reference model. Rebuild with tools/build_template.py. */\nwindow.CANVAS_MODEL=" + encoded + ";\n", encoding="utf-8")
    print(json.dumps({"output": str(args.output), "bytes": args.output.stat().st_size, "stats": model["stats"], "anomalies": model["anomalies"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
