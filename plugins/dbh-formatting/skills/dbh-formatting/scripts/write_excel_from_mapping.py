from __future__ import annotations

import argparse
import html
import json
import re
import shutil
import time
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET


CELL_RE = re.compile(rb"<c\b[^>]*?(?:/>|>.*?</c>)", re.DOTALL)
REL_RE = re.compile(
    rb"<Relationship\b[^>]*\bId=(?:\"([^\"]+)\"|'([^']+)')[^>]*\bTarget=(?:\"([^\"]+)\"|'([^']+)')"
)


def excel_serial(value: date) -> int:
    return (value - date(1899, 12, 30)).days


def xml_text(value: str) -> bytes:
    escaped = html.escape(value, quote=False)
    preserve = b' xml:space="preserve"' if value.startswith(" ") or value.endswith(" ") or "\n" in value else b""
    return b"<is><t" + preserve + b">" + escaped.encode("utf-8") + b"</t></is>"


def make_cell(address: str, value: str | int | None, attrs: bytes) -> bytes:
    attrs = re.sub(rb"\s+t=(?:\"[^\"]*\"|'[^']*')", b"", attrs)
    attrs = re.sub(rb"\s+r=(?:\"[^\"]*\"|'[^']*')", b"", attrs)
    head = b'<c r="' + address.encode("ascii") + b'"' + attrs
    if value is None:
        return head + b"/>"
    if isinstance(value, int):
        return head + b"><v>" + str(value).encode("ascii") + b"</v></c>"
    return head + b' t="inlineStr">' + xml_text(value) + b"</c>"


def replace_cell(xml: bytes, address: str, value: str | int | None, style_index: int | None = None) -> bytes:
    found = False

    def replace(match: re.Match[bytes]) -> bytes:
        nonlocal found
        address_match = re.search(rb"\br=(?:\"([^\"]+)\"|'([^']+)')", match.group(0))
        actual = (address_match.group(1) or address_match.group(2)).decode("utf-8") if address_match else ""
        if actual != address:
            return match.group(0)
        found = True
        raw = match.group(0)
        close = raw.find(b">")
        attrs = raw[2:close].rstrip().rstrip(b"/")
        if style_index is not None:
            attrs = re.sub(rb"\s+s=(?:\"[^\"]*\"|'[^']*')", b"", attrs)
            attrs += b' s="' + str(style_index).encode("ascii") + b'"'
        return make_cell(address, value, attrs)

    updated = CELL_RE.sub(replace, xml)
    if not found:
        raise KeyError(address)
    return updated


def expand_range(spec: str) -> list[str]:
    match = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", spec.upper())
    if not match:
        return [spec]
    c1, r1, c2, r2 = match.groups()

    def col_number(col: str) -> int:
        value = 0
        for char in col:
            value = value * 26 + ord(char) - 64
        return value

    def col_name(number: int) -> str:
        result = ""
        while number:
            number, rem = divmod(number - 1, 26)
            result = chr(65 + rem) + result
        return result

    return [f"{col_name(col)}{row}" for row in range(int(r1), int(r2) + 1) for col in range(col_number(c1), col_number(c2) + 1)]


def resolve_sheet_path(entries: dict[str, bytes], sheet_name: str | None) -> tuple[str, str]:
    workbook = entries["xl/workbook.xml"]
    rels = entries["xl/_rels/workbook.xml.rels"]
    sheets_match = re.search(rb"<sheets>(.*?)</sheets>", workbook, re.DOTALL)
    if not sheets_match:
        raise RuntimeError("workbook.xml has no sheets")
    candidates = list(re.finditer(rb"<sheet\b([^>]*)/?>", sheets_match.group(1)))
    if not candidates:
        raise RuntimeError("workbook.xml has no sheet entries")
    selected = candidates[0]
    selected_name = ""
    for candidate in candidates:
        attrs = candidate.group(1)
        name_match = re.search(rb"\bname=(?:\"([^\"]+)\"|'([^']+)')", attrs)
        name = (name_match.group(1) or name_match.group(2)).decode("utf-8") if name_match else ""
        if sheet_name is None or name == sheet_name:
            selected = candidate
            selected_name = name
            break
    if not selected_name:
        attrs = selected.group(1)
        name_match = re.search(rb"\bname=(?:\"([^\"]+)\"|'([^']+)')", attrs)
        selected_name = (name_match.group(1) or name_match.group(2)).decode("utf-8") if name_match else ""
    rid_match = re.search(rb"\br:id=(?:\"([^\"]+)\"|'([^']+)')", selected.group(1))
    if not rid_match:
        raise RuntimeError("Could not resolve worksheet relationship id")
    rid = (rid_match.group(1) or rid_match.group(2)).decode("utf-8")
    for rel in REL_RE.finditer(rels):
        rel_id = (rel.group(1) or rel.group(2)).decode("utf-8")
        if rel_id == rid:
            target = (rel.group(3) or rel.group(4)).decode("utf-8").lstrip("/").replace("../", "")
            return ("xl/" + target if not target.startswith("xl/") else target), selected_name
    raise RuntimeError(f"Could not resolve relationship {rid}")


def normalize(value: object, value_type: str | None, replace_fl_with_g: bool) -> str | int | None:
    if value is None:
        return None
    if value_type == "date":
        return excel_serial(date.fromisoformat(str(value)))
    if value_type == "integer":
        return int(value)
    result = str(value)
    if replace_fl_with_g:
        result = result.replace("FL", "G")
    return result


def parse_date_from_filename(path_value: object) -> date:
    """Read an ISO-style date from an approved customer-email file name only."""
    file_name = Path(str(path_value)).name
    patterns = (
        r"(?<!\d)(20\d{2})[._-](\d{1,2})[._-](\d{1,2})(?!\d)",
        r"(?<!\d)(20\d{2})(\d{2})(\d{2})(?!\d)",
    )
    for pattern in patterns:
        match = re.search(pattern, file_name)
        if match:
            year, month, day = (int(value) for value in match.groups())
            return date(year, month, day)
    raise ValueError(f"Không tìm thấy ngày YYYY.MM.DD, YYYY-MM-DD hoặc YYYYMMDD trong tên file email: {file_name}")


def parse_common_email_date(path_values: object) -> tuple[date, list[str]]:
    """Accept one or more approved email filenames only when their dates agree."""
    if not isinstance(path_values, list) or not path_values:
        raise ValueError("Cần ít nhất một đường dẫn email trong email_pdfs")
    file_paths = [str(value) for value in path_values if str(value).strip()]
    if not file_paths:
        raise ValueError("Cần ít nhất một đường dẫn email trong email_pdfs")
    parsed = [(parse_date_from_filename(path), Path(path).name) for path in file_paths]
    distinct_dates = {value for value, _ in parsed}
    if len(distinct_dates) != 1:
        detail = ", ".join(f"{name}: {value.isoformat()}" for value, name in parsed)
        raise ValueError(f"Nhiều file email có ngày khác nhau: {detail}")
    return parsed[0][0], [name for _, name in parsed]


def normalize_fc_floor(value: object, replace_fl_with_g: bool) -> str | int | None:
    """Normalize floor notation used only in Fc bands for the confirmed template."""
    normalized = normalize(value, "string", replace_fl_with_g)
    if isinstance(normalized, str) and normalized.strip().upper() in {"R", "RSL"}:
        return "RG"
    return normalized


def cell_style_index(worksheet: bytes, address: str) -> int:
    for match in CELL_RE.finditer(worksheet):
        address_match = re.search(rb"\br=(?:\"([^\"]+)\"|'([^']+)')", match.group(0))
        actual = (address_match.group(1) or address_match.group(2)).decode("utf-8") if address_match else ""
        if actual != address:
            continue
        style_match = re.search(rb"\bs=(?:\"(\d+)\"|'(\d+)')", match.group(0))
        return int(style_match.group(1) or style_match.group(2)) if style_match else 0
    raise KeyError(address)


def with_count(tag: bytes, count: int) -> bytes:
    if re.search(rb"\bcount=(?:\"\d+\"|'\d+')", tag):
        return re.sub(rb"\bcount=(?:\"\d+\"|'\d+')", f'count="{count}"'.encode("ascii"), tag, count=1)
    return tag[:-1] + f' count="{count}">'.encode("ascii")


def font_with_color(font: bytes, color: str) -> bytes:
    color_tag = f'<color rgb="{color}"/>'.encode("ascii")
    color_pattern = rb"<color\b[^>]*?(?:/>|>.*?</color>)"
    if re.search(color_pattern, font, flags=re.DOTALL):
        return re.sub(color_pattern, color_tag, font, count=1, flags=re.DOTALL)
    if font.endswith(b"/>"):
        return font[:-2] + b">" + color_tag + b"</font>"
    return font.replace(b"</font>", color_tag + b"</font>", 1)


def color_cell_styles(styles: bytes, base_and_color: dict[str, tuple[int, str]]) -> tuple[bytes, dict[str, int]]:
    """Clone existing cell styles with a red font, preserving all other formatting."""
    font_block = re.search(rb"(<fonts\b[^>]*>)(.*?)(</fonts>)", styles, flags=re.DOTALL)
    xf_block = re.search(rb"(<cellXfs\b[^>]*>)(.*?)(</cellXfs>)", styles, flags=re.DOTALL)
    if not font_block or not xf_block:
        raise RuntimeError("styles.xml thiếu fonts hoặc cellXfs")

    fonts = list(re.finditer(rb"<font\b[^>]*?(?:/>|>.*?</font>)", font_block.group(2), flags=re.DOTALL))
    xfs = list(re.finditer(rb"<xf\b[^>]*?(?:/>|>.*?</xf>)", xf_block.group(2), flags=re.DOTALL))
    if not fonts or not xfs:
        raise RuntimeError("styles.xml không có font hoặc cell style để sao chép")

    pair_to_new_style: dict[tuple[int, str], int] = {}
    new_fonts: list[bytes] = []
    new_xfs: list[bytes] = []
    for base_style, color in dict.fromkeys(base_and_color.values()):
        if base_style >= len(xfs):
            raise RuntimeError(f"Không tìm thấy cell style {base_style} trong styles.xml")
        base_xf = xfs[base_style].group(0)
        font_match = re.search(rb"\bfontId=(?:\"(\d+)\"|'(\d+)')", base_xf)
        font_id = int(font_match.group(1) or font_match.group(2)) if font_match else 0
        if font_id >= len(fonts):
            raise RuntimeError(f"Không tìm thấy font {font_id} trong styles.xml")
        new_font_id = len(fonts) + len(new_fonts)
        new_style_id = len(xfs) + len(new_xfs)
        new_fonts.append(font_with_color(fonts[font_id].group(0), color))
        if font_match:
            new_xf = re.sub(rb"\bfontId=(?:\"\d+\"|'\d+')", f'fontId="{new_font_id}"'.encode("ascii"), base_xf, count=1)
        else:
            new_xf = base_xf[:-2] + f' fontId="{new_font_id}"/>'.encode("ascii") if base_xf.endswith(b"/>") else base_xf.replace(b"<xf", f'<xf fontId="{new_font_id}"'.encode("ascii"), 1)
        new_xfs.append(new_xf)
        pair_to_new_style[(base_style, color)] = new_style_id

    fonts_open = with_count(font_block.group(1), len(fonts) + len(new_fonts))
    styles = styles[:font_block.start()] + fonts_open + font_block.group(2) + b"".join(new_fonts) + font_block.group(3) + styles[font_block.end():]
    xf_block = re.search(rb"(<cellXfs\b[^>]*>)(.*?)(</cellXfs>)", styles, flags=re.DOTALL)
    assert xf_block is not None
    xfs_open = with_count(xf_block.group(1), len(xfs) + len(new_xfs))
    styles = styles[:xf_block.start()] + xfs_open + xf_block.group(2) + b"".join(new_xfs) + xf_block.group(3) + styles[xf_block.end():]
    return styles, {address: pair_to_new_style[pair] for address, pair in base_and_color.items()}


def main() -> None:
    parser = argparse.ArgumentParser(description="Patch a copied Excel template from project data JSON and mapping config.")
    parser.add_argument("--config", required=True)
    args = parser.parse_args()
    started = time.perf_counter()
    config_path = Path(args.config).resolve()
    config = json.loads(config_path.read_text(encoding="utf-8"))

    def resolve_path(value: str) -> Path:
        path = Path(value)
        return path if path.is_absolute() else config_path.parent / path

    data_path = resolve_path(config["data_json"])
    data = json.loads(data_path.read_text(encoding="utf-8"))
    values = data.get("values", {})
    source = resolve_path(config["template"])
    output = resolve_path(config["output"])
    report_path = resolve_path(config["report"])
    replace_fl = bool(config.get("replace_FL_with_G", True))

    with zipfile.ZipFile(source, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
    sheet_path, actual_sheet = resolve_sheet_path(entries, config.get("sheet_name"))
    worksheet = entries[sheet_path]
    mapping: dict[str, str | int | None] = {}
    review_items: list[dict[str, str]] = []
    red_cell_colors: dict[str, str] = {}
    derived_email_dates: dict[str, str] | None = None
    for item in data.get("review_items", []):
        if isinstance(item, dict):
            review_items.append({str(key): str(value) for key, value in item.items()})

    for address, rule in config.get("cells", {}).items():
        source_key = rule.get("source")
        if source_key not in values:
            review_items.append({"field": address, "status": "missing_data", "note": f"data_json thiếu key: {source_key}"})
            continue
        mapping[address] = normalize(values[source_key], rule.get("type"), replace_fl)

    for spec in config.get("clear_cells", []) + config.get("clear_ranges", []):
        for address in expand_range(spec):
            mapping[address] = None

    email_dates = config.get("email_dates")
    if email_dates:
        received_cell = str(email_dates.get("received_cell", "D7"))
        return_cell = str(email_dates.get("return_due_cell", "H7"))
        unresolved_text = str(email_dates.get("unresolved_text", "要確認"))
        unresolved_color = str(email_dates.get("unresolved_font_color", "FFFF0000")).lstrip("#").upper()
        try:
            if email_dates.get("email_pdf"):
                email_paths = [email_dates["email_pdf"]]
            else:
                email_paths = email_dates.get("email_pdfs")
            received_date, file_names = parse_common_email_date(email_paths)
            return_date = received_date + timedelta(days=int(email_dates.get("return_days", 10)))
            mapping[received_cell] = excel_serial(received_date)
            mapping[return_cell] = excel_serial(return_date)
            derived_email_dates = {
                "source": "email file name",
                "file_names": file_names,
                "received_date": received_date.isoformat(),
                "return_due_date": return_date.isoformat(),
            }
        except (KeyError, TypeError, ValueError) as exc:
            mapping[received_cell] = unresolved_text
            mapping[return_cell] = unresolved_text
            red_cell_colors[received_cell] = unresolved_color
            red_cell_colors[return_cell] = unresolved_color
            review_items.append({"field": "依頼受理日・返却予定日", "status": "unclear", "note": str(exc)})

    scope_review = config.get("scope_review")
    if scope_review:
        scope_cell = str(scope_review.get("cell", "C16"))
        source_key = str(scope_review.get("source", "scope"))
        source_cell = scope_review.get("source_cell")
        citation_key = str(scope_review.get("citation_source", "scope_source"))
        unresolved_text = str(scope_review.get("unresolved_text", "要確認"))
        unresolved_color = str(scope_review.get("unresolved_font_color", "FFFF0000")).lstrip("#").upper()
        scope_value = values.get(source_key)
        if scope_value is None or not str(scope_value).strip():
            mapping[scope_cell] = unresolved_text
            if source_cell:
                mapping[str(source_cell)] = None
            red_cell_colors[scope_cell] = unresolved_color
            review_items.append({
                "field": str(scope_review.get("field", "検討部位／検討範囲")),
                "status": "unclear",
                "note": str(scope_review.get("note", "構造図内に対象範囲の明記なし。")),
            })
        else:
            normalized_scope = normalize(scope_value, scope_review.get("type", "string"), replace_fl)
            mapping[scope_cell] = normalized_scope
            if source_cell:
                scope_source = values.get(citation_key)
                mapping[str(source_cell)] = normalize(scope_source, "string", replace_fl) if scope_source is not None and str(scope_source).strip() else None
            if normalized_scope == unresolved_text:
                red_cell_colors[scope_cell] = unresolved_color

    review_notes = config.get("review_notes")
    if review_notes:
        statuses = {str(value) for value in review_notes.get("include_statuses", [])}
        candidates = [
            item for item in review_items
            if not statuses or str(item.get("status", "")) in statuses
        ]
        cells = list(review_notes.get("cells", []))
        template = review_notes.get("template", "・【要確認】{field}：{note}")
        color = str(review_notes.get("font_color", "FFFF0000")).lstrip("#").upper()
        for address, item in zip(cells, candidates):
            mapping[address] = template.format(
                field=item.get("field", ""),
                status=item.get("status", ""),
                note=item.get("note", ""),
            )
            red_cell_colors[address] = color
        if len(candidates) > len(cells):
            review_items.append({
                "field": "review_notes",
                "status": "overflow",
                "note": f"{len(candidates)} mục cần ghi chú nhưng chỉ có {len(cells)} ô ghi chú",
            })

    fc_config = config.get("fc_table")
    if fc_config:
        fc_source = data.get("fc_table", {})
        fc_rows = fc_source.get("rows", [])
        slots: list[list[str]] = []
        for block in fc_config.get("blocks", []):
            columns = block["columns"]
            for row in block["rows"]:
                slots.append([f"{column}{row}" for column in columns])
        if fc_config.get("label_cell") and fc_source.get("source_label"):
            template = fc_config.get("label_template", "Fc\n({source_label})")
            mapping[fc_config["label_cell"]] = template.format(source_label=fc_source["source_label"])
        if len(fc_rows) > len(slots):
            review_items.append({"field": "fc_table", "status": "overflow", "note": f"{len(fc_rows)} dòng Fc nhưng template chỉ có {len(slots)} slot"})
        for index, row_data in enumerate(fc_rows[: len(slots)]):
            cells = slots[index]
            mapping[cells[0]] = normalize(row_data.get("fc"), "integer", replace_fl)
            mapping[cells[1]] = normalize_fc_floor(row_data.get("from"), replace_fl)
            mapping[cells[2]] = row_data.get("separator", "～")
            mapping[cells[3]] = normalize_fc_floor(row_data.get("to"), replace_fl)
        if fc_rows and not fc_source.get("source_label"):
            review_items.append({"field": "fc_table.source_label", "status": "missing_data", "note": "Chưa có trang nguồn Fc"})

    citation_items: list[dict[str, str]] = []
    for item in data.get("citations", []):
        if isinstance(item, dict):
            citation_items.append({str(key): str(value) for key, value in item.items()})
    citation_config = config.get("citation_notes")
    if citation_config:
        cells = list(citation_config.get("cells", []))
        template = citation_config.get("template", "・参照：{text}")
        for address, item in zip(cells, citation_items):
            mapping[address] = template.format(**item)
        if len(citation_items) > len(cells):
            review_items.append({
                "field": "citation_notes",
                "status": "overflow",
                "note": f"{len(citation_items)} mục trích dẫn nhưng chỉ có {len(cells)} ô trích dẫn",
            })

    for address, value in mapping.items():
        if isinstance(value, str) and "要確認" in value:
            red_cell_colors.setdefault(address, "FFFF0000")

    colored_styles: dict[str, int] = {}
    if red_cell_colors:
        style_inputs = {
            address: (cell_style_index(worksheet, address), color)
            for address, color in red_cell_colors.items()
            if mapping.get(address) not in (None, "")
        }
        if style_inputs:
            styles_path = "xl/styles.xml"
            if styles_path not in entries:
                raise RuntimeError("Workbook không có xl/styles.xml để tô đỏ các mục 要確認")
            entries[styles_path], colored_styles = color_cell_styles(entries[styles_path], style_inputs)

    for address, value in mapping.items():
        try:
            worksheet = replace_cell(worksheet, address, value, colored_styles.get(address))
        except KeyError:
            review_items.append({"field": address, "status": "missing_cell", "note": "Ô đích không tồn tại trong worksheet XML"})
    entries[sheet_path] = worksheet

    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, output)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, content in entries.items():
            zout.writestr(name, content)

    with zipfile.ZipFile(output, "r") as zout:
        zip_error = zout.testzip()
        xml_errors = []
        for name in zout.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                try:
                    ET.fromstring(zout.read(name))
                except Exception as exc:
                    xml_errors.append({"part": name, "error": str(exc)})

    controls = lambda names: sorted(name for name in names if any(k in name.lower() for k in ("ctrlprop", "vml", "drawing", "control")))
    with zipfile.ZipFile(source, "r") as original_zip, zipfile.ZipFile(output, "r") as output_zip:
        original_controls = controls(original_zip.namelist())
        output_controls = controls(output_zip.namelist())

    report = {
        "schema": "dbh-formatting/excel-mapping-report@0.1",
        "generated_at": datetime.now().isoformat(),
        "source_template": str(source),
        "data_json": str(data_path),
        "output_workbook": str(output),
        "worksheet": actual_sheet,
        "mapping_count": len(mapping),
        "citation_items": citation_items,
        "derived_email_dates": derived_email_dates,
        "red_review_cells": sorted(colored_styles),
        "review_items": review_items,
        "validation": {
            "zip_test": zip_error,
            "xml_errors": xml_errors,
            "control_parts_original": len(original_controls),
            "control_parts_output": len(output_controls),
            "control_parts_preserved": original_controls == output_controls,
        },
        "timing_ms": round((time.perf_counter() - started) * 1000, 1),
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()

