import json
import math
import os
import posixpath
import sys
import zipfile
from copy import copy
from pathlib import Path
from xml.etree import ElementTree

from openpyxl import load_workbook

from cra_support import parse_cra_paste_text
from roll_forward_core import (
    RollForwardWarnings,
    SubjectConfig,
    find_expense_bkd_header_row,
    find_expense_bkd_total_row,
    find_header_col,
    find_header_col_near,
    find_header_row,
    find_prior_file,
    find_q1_bkd_header_row,
    normalize_text,
    process_lead_sheet,
)


ROOT = Path(__file__).resolve().parent.parent
SOURCE = Path(__file__).resolve().parent
TEMPLATES = ROOT / "templates"
PRIOR_DIR = ROOT / "常规样本" / "吉安"
OUTPUT_DIR = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else ROOT / "tmp_live_integrated_corefix3"

FAILURES = []
PASSES = []


def check(condition, message):
    if condition:
        PASSES.append(message)
    else:
        FAILURES.append(message)


def workbook_path(prefix):
    matches = sorted(OUTPUT_DIR.glob(f"{prefix}*.xlsx"))
    if not matches:
        raise FileNotFoundError(f"Missing output workbook: {prefix}")
    return matches[-1]


def sheet_name(wb, prefix):
    return next(name for name in wb.sheetnames if name.startswith(prefix))


def is_roll_highlight(cell):
    return (
        cell.fill.fill_type == "solid"
        and cell.fill.fgColor.type == "rgb"
        and str(cell.fill.fgColor.rgb).upper().endswith("FFFF99")
    )


def has_overlapping_merges(ws):
    ranges = list(ws.merged_cells.ranges)
    for index, first in enumerate(ranges):
        for second in ranges[index + 1 :]:
            separated = (
                first.max_row < second.min_row
                or second.max_row < first.min_row
                or first.max_col < second.min_col
                or second.max_col < first.min_col
            )
            if not separated:
                return True
    return False


def relationship_source(rels_name):
    if rels_name == "_rels/.rels":
        return ""
    return posixpath.join(
        posixpath.dirname(posixpath.dirname(rels_name)),
        posixpath.basename(rels_name)[:-5],
    )


def check_xlsx_package(path):
    errors = []
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        if archive.testzip() is not None:
            errors.append("CRC failure")
        for name in names:
            if not name.endswith((".xml", ".rels")):
                continue
            try:
                root = ElementTree.fromstring(archive.read(name))
            except Exception as exc:
                errors.append(f"Invalid XML {name}: {exc}")
                continue
            if not name.endswith(".rels"):
                continue
            source = relationship_source(name)
            base = posixpath.dirname(source)
            for rel in root:
                if rel.attrib.get("TargetMode") == "External":
                    continue
                target = rel.attrib.get("Target", "")
                resolved = (
                    target.lstrip("/")
                    if target.startswith("/")
                    else posixpath.normpath(posixpath.join(base, target))
                )
                if resolved not in names:
                    errors.append(f"Missing relationship target {name} -> {resolved}")
    wb = load_workbook(path, data_only=False)
    for ws in wb.worksheets:
        if has_overlapping_merges(ws):
            errors.append(f"Overlapping merges: {ws.title}")
        for validation in ws.data_validations.dataValidation:
            try:
                str(validation.sqref)
            except Exception as exc:
                errors.append(f"Invalid validation range {ws.title}: {exc}")
    wb.close()
    check(not errors, f"{path.name}: ZIP/XML/relationships/merge ranges clean")
    if errors:
        FAILURES.extend(f"{path.name}: {error}" for error in errors)


def numeric_equal(left, right):
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return math.isclose(float(left), float(right), rel_tol=1e-9, abs_tol=0.01)
    return left == right


def expense_records(ws, value_kind):
    header = find_expense_bkd_header_row(ws)
    total = find_expense_bkd_total_row(ws, header)
    if value_kind == "current":
        value_col = find_header_col(ws, header, ["本期账面审定数", "本期审定数", "本期数"])
    else:
        value_col = find_header_col(ws, header, ["上期末审定数", "上期审定数", "上年数", "PY"])
    if not value_col:
        keywords = ["本期账面审定数", "本期审定数"] if value_kind == "current" else ["上期末审定数", "上期审定数", "PY"]
        value_col = find_header_col_near(ws, header, keywords, row_offsets=(-1, 0))
    descriptor_cols = []
    for keywords in (["账套名称/账套编码", "账套名称"], ["科目编码"], ["科目名称"]):
        col = find_header_col(ws, header, keywords)
        if col:
            descriptor_cols.append(col)
    records = []
    for row in range(header + 1, total):
        descriptor = tuple(normalize_text(ws.cell(row, col).value) for col in descriptor_cols)
        if any(descriptor):
            records.append((descriptor, ws.cell(row, value_col).value))
    return records


def check_expense_roll(prior_path, output_path, sheet_prefix):
    prior_wb = load_workbook(prior_path, data_only=True)
    output_wb = load_workbook(output_path, data_only=False)
    prior_records = expense_records(prior_wb[sheet_name(prior_wb, sheet_prefix)], "current")
    output_records = expense_records(output_wb[sheet_name(output_wb, sheet_prefix)], "py")
    prior_wb.close()
    output_wb.close()
    same = len(prior_records) == len(output_records) and all(
        p_desc == o_desc and numeric_equal(p_value, o_value)
        for (p_desc, p_value), (o_desc, o_value) in zip(prior_records, output_records)
    )
    check(same, f"{sheet_prefix}: prior current values rolled to current PY without row loss")


def check_q1(path):
    wb = load_workbook(path, data_only=False)
    q101 = wb[sheet_name(wb, "Q1.01")]
    header = find_q1_bkd_header_row(q101)
    total = next(
        row
        for row in range(header + 1, q101.max_row + 1)
        if any(normalize_text(q101.cell(row, col).value) == "合计" for col in range(1, 8))
    )
    right_side_ok = all(
        q101.cell(row, 15).value in (None, "")
        and q101.cell(row, 17).value in (None, "")
        and q101.cell(row, 18).value in (None, "")
        and isinstance(q101.cell(row, 16).value, str)
        and q101.cell(row, 16).value.startswith("=")
        for row in range(header + 1, total)
    )
    check(right_side_ok, "Q1.01: current/non-current/covenant/Notes1 prior values are not rolled")
    for prefix in ("Q1.02b", "Q1.02c"):
        ws = wb[sheet_name(wb, prefix)]
        count = sum(is_roll_highlight(cell) for row in ws.iter_rows() for cell in row)
        check(count == 0, f"{prefix}: no roll-forward yellow added")
    q105 = wb[sheet_name(wb, "Q1.05")]
    fixed_labels = ("B37", "B38", "B39", "B41", "B43", "B48", "B51", "B55", "B59", "B60", "B61", "B63", "B65", "B70", "B73", "B77")
    check(all(not is_roll_highlight(q105[cell]) for cell in fixed_labels), "Q1.05: template labels are not marked yellow")
    check(not has_overlapping_merges(q105), "Q1.05: no overlapping merged cells")
    template_wb = load_workbook(TEMPLATES / "Q1 SWP 银行借款 202YMMDD XYZ公司.xlsx", data_only=False)
    template_q105 = template_wb[sheet_name(template_wb, "Q1.05")]
    check(
        all(q105.column_dimensions[col].width == template_q105.column_dimensions[col].width for col in "BCDEFGHI"),
        "Q1.05: current-template column widths retained",
    )
    template_wb.close()
    wb.close()

    drawing_rows = []
    media = []
    ns = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
    with zipfile.ZipFile(path) as archive:
        media = [name for name in archive.namelist() if name.startswith("xl/media/")]
        for name in archive.namelist():
            if not name.startswith("xl/drawings/drawing") or not name.endswith(".xml"):
                continue
            root = ElementTree.fromstring(archive.read(name))
            for anchor in list(root):
                marker = anchor.find("xdr:from", ns)
                if marker is not None and anchor.find("xdr:pic", ns) is not None:
                    drawing_rows.append(int(marker.find("xdr:row", ns).text))
    check(len(media) == 7, "Q1: exactly two Note images and five Q1.05 evidence images copied")
    check(sorted(drawing_rows) == [34, 34, 42, 42, 42, 64, 64], "Q1: copied images follow rebuilt Note blocks")


def thin_outer_border(ws, cell_range):
    cells = ws[cell_range]
    return (
        all(cell.border.top.style == "thin" for cell in cells[0])
        and all(cell.border.bottom.style == "thin" for cell in cells[-1])
        and all(row[0].border.left.style == "thin" for row in cells)
        and all(row[-1].border.right.style == "thin" for row in cells)
    )


def check_uexp(path):
    wb = load_workbook(path, data_only=False)
    lead = wb["Uexp_Lead"]
    bkd = wb["Uexp财务费用BKD"]
    for ws, ranges in (
        (lead, ("C13:G15", "C35:G37")),
        (bkd, ("C19:G21", "C46:G46", "C47:G49", "C50:G51", "C52:G52")),
    ):
        merged = {str(item) for item in ws.merged_cells.ranges}
        for cell_range in ranges:
            check(cell_range in merged, f"Uexp {ws.title} {cell_range}: merged wording box retained")
            check(thin_outer_border(ws, cell_range), f"Uexp {ws.title} {cell_range}: thin continuous outer border")
            check(ws[cell_range.split(":")[0]].alignment.wrap_text is True, f"Uexp {ws.title} {cell_range}: wrap text enabled")
    wb.close()


def check_l2_expansion_fixture():
    template_path = TEMPLATES / "L2 SWP 长期待摊费用 202YMMDD XYZ公司.xlsx"
    prior_path = PRIOR_DIR / "L2 SWP 长期待摊费用 20251231吉安.xlsx"
    prior_formula = load_workbook(prior_path, data_only=False)
    prior_values = load_workbook(prior_path, data_only=True)
    output_wb = load_workbook(template_path, data_only=False)
    for workbook in (prior_formula, prior_values):
        ws = workbook["L2.00 Lead"]
        header = find_header_row(ws, "期末审定数", (1, 80))
        closing_col = find_header_col(ws, header, ["期末审定数"])
        ws.insert_rows(header + 3, 5)
        for offset in range(5):
            row = header + 3 + offset
            ws.cell(row, 2, f"Fixture company {offset + 3}")
            ws.cell(row, 3, f"18{offset + 3:02d}")
            ws.cell(row, 4, f"Fixture item {offset + 3}")
            ws.cell(row, closing_col, 1000 + offset)
    config = SubjectConfig(SOURCE / "subjects_config.json").get_subject("L2")["lead_sheet"]
    process_lead_sheet(
        prior_formula["L2.00 Lead"],
        output_wb["L2.00 Lead"],
        prior_values["L2.00 Lead"],
        {"PM": 1, "TE": 1, "SAD": 1},
        "2026-06-30",
        "Fixture company",
        config,
        RollForwardWarnings(),
    )
    ws = output_wb["L2.00 Lead"]
    header = find_header_row(ws, "期末审定数", (1, 80))
    total = next(
        row
        for row in range(header + 1, ws.max_row + 1)
        if any(normalize_text(ws.cell(row, col).value) == "合计" for col in range(1, 8))
    )
    detail_start = header + 1
    detail_end = detail_start + 7 - 1
    check(total == detail_end + 1, "L2 fixture: Lead total row moves below seven detail rows")
    check(
        all(ws.cell(row, 10).value == f"=G{row}+I{row}" and ws.cell(row, 13).value == f"=J{row}+L{row}" for row in range(detail_start, detail_end + 1)),
        "L2 fixture: every expanded detail row retains J/M formulas",
    )
    check(ws.cell(total, 7).value == f"=SUM(G{detail_start}:G{detail_end})", "L2 fixture: total formula uses expanded detail range")
    below_text = [normalize_text(ws.cell(row, col).value) for row in range(total + 1, min(total + 12, ws.max_row) + 1) for col in range(1, 14)]
    check("Rx" in below_text and "A3" in below_text and "Diff" in below_text, "L2 fixture: Rx/A3/Diff area moves below total")
    prior_formula.close()
    prior_values.close()
    output_wb.close()


def check_repair_roots(l2_path, vcvd_path):
    with zipfile.ZipFile(vcvd_path) as archive:
        external_link_parts = [name for name in archive.namelist() if "externalLinks/" in name]
    check(not external_link_parts, "VCVD: unstable external-link cache is not retained")

    wb = load_workbook(l2_path, data_only=False)
    ws = wb["L2.01.1 BKD"]
    header = find_header_row(ws, "项目编码", (1, 80))
    total = next(
        row
        for row in range(header + 1, ws.max_row + 1)
        if any(normalize_text(ws.cell(row, col).value) == "合计" for col in range(1, 8))
    )
    hidden_formula_rows = [
        row
        for row in range(28, total)
        if any(
            isinstance(ws.cell(row, col).value, str) and ws.cell(row, col).value.startswith("=")
            for col in range(19, ws.max_column + 1)
        )
    ]
    check(not hidden_formula_rows, "L2.01.1: expanded records do not clone hidden plug-in formulas")
    check(ws.cell(total, 7).value == f"=SUM(G{header + 1}:G{total - 1})", "L2.01.1: total starts at first detail row")
    wb.close()


def check_cra_applicable_variants():
    header = "科目\t认定\tCRA\t比例\t是否适用"
    for raw in ("N", "N/A", "n/a", "no", "否", "不适用"):
        records = parse_cra_paste_text(
            header + f"\n管理费用\t发生\tModerate\t50%\t{raw}",
            selected_subjects=["UexpVCVD"],
        )
        record = records[0]
        check(record["applicable"] is False and record["cra_level"] == "N/A" and record["ratio_text"] == "N/A", f"CRA applicable={raw}: writes N/A")
    for text, label in (
        (header + "\n管理费用\t发生\tModerate\t50%\t", "blank applicable"),
        ("科目\t认定\tCRA\t比例\n管理费用\t发生\tModerate\t50%", "missing applicable column"),
    ):
        record = parse_cra_paste_text(text, selected_subjects=["UexpVCVD"])[0]
        check(record["applicable"] is True and record["cra_level"] == "Moderate", f"CRA {label}: risk level retained")


def main():
    config = SubjectConfig(SOURCE / "subjects_config.json")
    matched_vcvd = find_prior_file(PRIOR_DIR, "UexpVCVD", "2025", config.get_subject("UexpVCVD"))
    matched_uexp = find_prior_file(PRIOR_DIR, "Uexp", "2025", config.get_subject("Uexp"))
    check(matched_vcvd and "VC&VD" in Path(matched_vcvd).name, "VCVD matcher selects the VC&VD prior workbook")
    check(matched_uexp and "财务费用" in Path(matched_uexp).name, "Uexp matcher selects the finance expense prior workbook")

    q1 = workbook_path("Q1 SWP")
    vcvd = workbook_path("U_exp SWP VC&VD")
    l2 = workbook_path("L2 SWP")
    uexp = workbook_path("U_exp SWP other")
    for path in (q1, vcvd, l2, uexp):
        check_xlsx_package(path)

    check_q1(q1)
    check_expense_roll(Path(matched_vcvd), vcvd, "VC.00")
    check_expense_roll(Path(matched_vcvd), vcvd, "VD.00")
    check_expense_roll(Path(matched_uexp), uexp, "Uexp财务费用BKD")
    check_uexp(uexp)
    check_l2_expansion_fixture()
    check_repair_roots(l2, vcvd)
    check_cra_applicable_variants()

    report = {"passes": len(PASSES), "failures": FAILURES, "checked_output": str(OUTPUT_DIR)}
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if FAILURES:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
