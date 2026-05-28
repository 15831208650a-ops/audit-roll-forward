#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
审计底稿 Roll Forward 核心模块 (v4.0)
功能：支持多科目差异的批量Roll Forward处理
作者：AI Assistant
"""

import os
import re
import shutil
import json
import datetime
import warnings
import sys
import fnmatch
from collections import OrderedDict
from copy import copy, deepcopy
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Border, PatternFill
from openpyxl.formula.translate import Translator
from openpyxl.formatting.formatting import ConditionalFormatting
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange

# 忽略openpyxl的警告
warnings.filterwarnings('ignore', category=UserWarning)


def resource_path(relative_path):
    """Return an absolute path for bundled resources in source or PyInstaller."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base_dir = Path(sys._MEIPASS)
    else:
        base_dir = Path(__file__).resolve().parent
    return str(base_dir / relative_path)

class RollForwardError(Exception):
    """Roll Forward 自定义异常"""
    pass


class SubjectConfig:
    """科目配置管理器"""

    def __init__(self, config_path=None):
        if config_path is None:
            config_path = resource_path("subjects_config.json")
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)

    def get_subject(self, subject_code):
        """获取单个科目配置"""
        return self.config.get("subjects", {}).get(subject_code)

    def get_all_subjects(self):
        """获取所有科目配置"""
        return self.config.get("subjects", {})

    def get_subject_list(self):
        """获取科目列表 [(code, name), ...]"""
        subjects = []
        for code, info in self.config.get("subjects", {}).items():
            subjects.append((code, info.get("name", "")))
        return subjects


def find_prior_file(prior_dir, subject_code, prior_year, config):
    """
    在上年底稿目录中查找对应科目的底稿文件
    支持多个搜索目录
    """
    # 支持多个搜索目录（主目录 + 子目录）
    search_dirs = [prior_dir]
    # 如果prior_dir不是根目录，也搜索根目录
    parent_dir = os.path.dirname(prior_dir)
    if parent_dir and parent_dir not in search_dirs:
        search_dirs.append(parent_dir)

    # 获取文件匹配模式
    patterns = config.get("prior_file_patterns")
    if not patterns:
        patterns = [config.get("prior_file_pattern", f"{subject_code}*")]

    for search_dir in search_dirs:
        if not os.path.exists(search_dir):
            continue

        candidates = [
            filename for filename in os.listdir(search_dir)
            if filename.endswith('.xlsx') and not filename.startswith('~')
        ]

        for pattern in patterns:
            concrete_pattern = pattern.replace('{prior_year}', str(prior_year))
            matched = [
                filename for filename in candidates
                if fnmatch.fnmatch(filename, concrete_pattern)
            ]
            if matched:
                matched.sort(
                    key=lambda filename: os.path.getmtime(os.path.join(search_dir, filename)),
                    reverse=True
                )
                return os.path.join(search_dir, matched[0])

        # 1. 尝试精确匹配：科目代码 + 年份
        for filename in candidates:
            if subject_code in filename and str(prior_year) in filename:
                return os.path.join(search_dir, filename)

        # 2. 放宽条件：只匹配科目代码
        for filename in candidates:
            if subject_code in filename:
                return os.path.join(search_dir, filename)

        # 3. 尝试用pattern中的关键字匹配（科目代码可能和文件名不完全一致）
        for pattern in patterns:
            for filename in candidates:
                # 提取pattern中的关键词（去掉通配符和年份）
                keywords = [k for k in pattern.replace('*', ' ').replace('{prior_year}', '').split() if k and k not in ['.xlsx', '']]
                if all(kw in filename for kw in keywords):
                    return os.path.join(search_dir, filename)

    return None


def find_header_row(ws, search_text, search_range=(1, 80)):
    """在工作表中查找包含特定文本的行"""
    start, end = search_range
    for row in range(start, min(end + 1, ws.max_row + 1)):
        for col in range(1, min(20, ws.max_column + 1)):
            cell_value = ws.cell(row=row, column=col).value
            if cell_value and search_text in str(cell_value):
                return row
    return None


def parse_date_value(value):
    """Parse common user-entered date formats."""
    if isinstance(value, datetime.datetime):
        return value
    if isinstance(value, datetime.date):
        return datetime.datetime(value.year, value.month, value.day)
    if value is None:
        return None

    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def normalize_text(value):
    """Normalize labels for cross-template matching."""
    if value is None:
        return ""
    text = str(value)
    text = re.sub(r"\s+", "", text)
    text = text.replace("（", "(").replace("）", ")")
    return text


def fill_adjacent_header_value(ws, keywords, value, search_range=(1, 20)):
    """Fill the cell to the right of a matching header label."""
    if value is None or str(value).strip() == "":
        return False

    start, end = search_range
    for row in range(start, min(end + 1, ws.max_row + 1)):
        for col in range(1, min(20, ws.max_column + 1)):
            cell_value = ws.cell(row=row, column=col).value
            if not cell_value:
                continue

            text = str(cell_value)
            if any(keyword in text for keyword in keywords):
                return set_cell_value(ws, row, col + 1, value)

    return False


def normalize_lead_date_formats(ws):
    """Keep report-date cells and date formulas in a consistent display format."""
    for row in range(1, min(80, ws.max_row) + 1):
        for col in range(1, min(20, ws.max_column) + 1):
            cell = ws.cell(row=row, column=col)
            value = cell.value
            if isinstance(value, (datetime.datetime, datetime.date)):
                cell.number_format = "yyyy/mm/dd"
            elif isinstance(value, str) and (
                "$C$3" in value
                or "DATE(YEAR(" in value
                or "汇总!D5" in value
                or "汇总!$D$5" in value
            ):
                cell.number_format = "yyyy/mm/dd"


def set_cell_value(ws, row, column, value):
    """Set a cell value unless the target is a merged placeholder."""
    cell = ws.cell(row=row, column=column)
    if isinstance(cell, MergedCell):
        return False
    cell.value = value
    return True


def find_header_col(ws, row, keywords, max_col=None):
    """Find a column in one header row by keyword."""
    if not row:
        return None
    max_col = max_col or ws.max_column
    for col in range(1, min(max_col, ws.max_column) + 1):
        value = ws.cell(row=row, column=col).value
        if not value:
            continue
        text = normalize_text(value)
        if any(normalize_text(keyword) in text for keyword in keywords):
            return col
    return None


def find_header_col_near(ws, header_row, keywords, row_offsets=(-1, 0, 1), max_col=None):
    """Find a header column in rows around the main header row."""
    for offset in row_offsets:
        row = header_row + offset
        if row < 1:
            continue
        col = find_header_col(ws, row, keywords, max_col)
        if col:
            return col
    return None


def row_contains_any(ws, row, keywords, columns=None):
    """Return True when any target column in a row contains a keyword."""
    if columns is None:
        columns = range(1, min(20, ws.max_column) + 1)
    row_text = " ".join(str(ws.cell(row=row, column=col).value or "") for col in columns)
    normalized = normalize_text(row_text)
    return any(normalize_text(keyword) in normalized for keyword in keywords)


def copy_cell_shape(source_cell, target_cell, translate_formula=False):
    """Copy style and optionally formula/value from one cell to another."""
    if isinstance(target_cell, MergedCell):
        return

    if source_cell.has_style:
        target_cell._style = copy(source_cell._style)
    if source_cell.number_format:
        target_cell.number_format = source_cell.number_format
    if source_cell.font:
        target_cell.font = copy(source_cell.font)
    if source_cell.fill:
        target_cell.fill = copy(source_cell.fill)
    if source_cell.border:
        target_cell.border = copy(source_cell.border)
    if source_cell.alignment:
        target_cell.alignment = copy(source_cell.alignment)
    if source_cell.protection:
        target_cell.protection = copy(source_cell.protection)

    value = source_cell.value
    if translate_formula and isinstance(value, str) and value.startswith("="):
        try:
            value = Translator(value, origin=source_cell.coordinate).translate_formula(target_cell.coordinate)
        except Exception:
            pass
    target_cell.value = value


def copy_row_shape(ws, source_row, target_row, translate_formula=True):
    """Copy a row's visible structure into another row."""
    for col in range(1, ws.max_column + 1):
        copy_cell_shape(
            ws.cell(row=source_row, column=col),
            ws.cell(row=target_row, column=col),
            translate_formula=translate_formula,
        )
    if source_row in ws.row_dimensions:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def transform_cell_range_for_insert(cell_range, idx, amount, axis="row", expand_adjacent=False):
    """Return a range adjusted for inserted rows or columns."""
    new_range = CellRange(str(cell_range))

    if axis == "row":
        if new_range.min_row >= idx:
            new_range.shift(row_shift=amount)
        elif new_range.min_row < idx <= new_range.max_row:
            new_range.expand(down=amount)
        elif expand_adjacent and new_range.max_row == idx - 1:
            new_range.expand(down=amount)
    else:
        if new_range.min_col >= idx:
            new_range.shift(col_shift=amount)
        elif new_range.min_col < idx <= new_range.max_col:
            new_range.expand(right=amount)
        elif expand_adjacent and new_range.max_col == idx - 1:
            new_range.expand(right=amount)

    return new_range


def transform_multi_cell_range_for_insert(sqref, idx, amount, axis="row", expand_adjacent=False):
    """Adjust all ranges in a MultiCellRange for an insertion."""
    if sqref is None:
        return sqref

    ranges = getattr(sqref, "ranges", None)
    if ranges is None:
        ranges = [CellRange(str(sqref))]

    return MultiCellRange([
        transform_cell_range_for_insert(cell_range, idx, amount, axis, expand_adjacent)
        for cell_range in ranges
    ])


def update_conditional_formatting_for_insert(ws, idx, amount, axis="row"):
    """Move and extend conditional formatting ranges around inserted rows/columns."""
    cf_rules = getattr(ws.conditional_formatting, "_cf_rules", None)
    if not cf_rules:
        return

    updated_rules = OrderedDict()
    for conditional_formatting, rules in cf_rules.items():
        new_sqref = transform_multi_cell_range_for_insert(
            conditional_formatting.sqref,
            idx,
            amount,
            axis=axis,
            expand_adjacent=True,
        )
        new_cf = ConditionalFormatting(
            sqref=new_sqref,
            pivot=conditional_formatting.pivot,
            cfRule=conditional_formatting.cfRule,
        )
        updated_rules[new_cf] = rules

    ws.conditional_formatting._cf_rules = updated_rules


def update_data_validations_for_insert(ws, idx, amount, axis="row"):
    """Move and extend data validation ranges around inserted rows/columns."""
    validations = getattr(ws.data_validations, "dataValidation", [])
    for validation in validations:
        validation.sqref = transform_multi_cell_range_for_insert(
            validation.sqref,
            idx,
            amount,
            axis=axis,
            expand_adjacent=True,
        )


def insert_rows_preserving_sheet_metadata(ws, idx, amount):
    """Insert rows without leaving sheet-level metadata at stale coordinates."""
    if amount <= 0:
        return

    affected_ranges = []
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row >= idx:
            affected_ranges.append((merged_range, "shift"))
        elif merged_range.min_row < idx <= merged_range.max_row:
            affected_ranges.append((merged_range, "expand"))

    ws.insert_rows(idx, amount)

    for merged_range, action in affected_ranges:
        if action == "shift":
            merged_range.shift(row_shift=amount)
        elif action == "expand":
            merged_range.expand(down=amount)

    update_conditional_formatting_for_insert(ws, idx, amount, axis="row")
    update_data_validations_for_insert(ws, idx, amount, axis="row")


def insert_cols_preserving_sheet_metadata(ws, idx, amount):
    """Insert columns without leaving sheet-level metadata at stale coordinates."""
    if amount <= 0:
        return

    affected_ranges = []
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_col >= idx:
            affected_ranges.append((merged_range, "shift"))
        elif merged_range.min_col < idx <= merged_range.max_col:
            affected_ranges.append((merged_range, "expand"))

    ws.insert_cols(idx, amount)

    for merged_range, action in affected_ranges:
        if action == "shift":
            merged_range.shift(col_shift=amount)
        elif action == "expand":
            merged_range.expand(right=amount)

    update_conditional_formatting_for_insert(ws, idx, amount, axis="col")
    update_data_validations_for_insert(ws, idx, amount, axis="col")


def clone_worksheet_contents(ws_source, ws_target):
    """Replace a worksheet's contents with another worksheet's contents."""
    for merged_range in list(ws_target.merged_cells.ranges):
        ws_target.unmerge_cells(str(merged_range))

    ws_target._cells = {}
    ws_target.merged_cells = deepcopy(ws_source.merged_cells)
    ws_target.sheet_format = copy(ws_source.sheet_format)
    ws_target.sheet_properties = copy(ws_source.sheet_properties)
    ws_target.page_margins = copy(ws_source.page_margins)
    ws_target.page_setup = copy(ws_source.page_setup)
    ws_target.print_options = copy(ws_source.print_options)
    ws_target.freeze_panes = ws_source.freeze_panes
    ws_target.auto_filter.ref = ws_source.auto_filter.ref

    ws_target.row_dimensions = deepcopy(ws_source.row_dimensions)
    ws_target.column_dimensions = deepcopy(ws_source.column_dimensions)
    ws_target.conditional_formatting = deepcopy(ws_source.conditional_formatting)
    ws_target.data_validations = deepcopy(ws_source.data_validations)
    ws_target._tables = deepcopy(ws_source._tables)

    for row in ws_source.iter_rows():
        for source_cell in row:
            target_cell = ws_target.cell(row=source_cell.row, column=source_cell.column)
            if source_cell.value is not None:
                target_cell.value = source_cell.value
            if source_cell.has_style:
                target_cell._style = copy(source_cell._style)
            if source_cell.number_format:
                target_cell.number_format = source_cell.number_format
            if source_cell.font:
                target_cell.font = copy(source_cell.font)
            if source_cell.fill:
                target_cell.fill = copy(source_cell.fill)
            if source_cell.border:
                target_cell.border = copy(source_cell.border)
            if source_cell.alignment:
                target_cell.alignment = copy(source_cell.alignment)
            if source_cell.protection:
                target_cell.protection = copy(source_cell.protection)
            if source_cell.hyperlink:
                target_cell._hyperlink = copy(source_cell.hyperlink)
            if source_cell.comment:
                target_cell.comment = copy(source_cell.comment)


def find_row_containing(ws, text, search_range=(1, 120)):
    """Find the first row containing a text fragment."""
    start, end = search_range
    for row in range(start, min(end, ws.max_row) + 1):
        for col in range(1, min(30, ws.max_column) + 1):
            value = ws.cell(row=row, column=col).value
            if value and text in str(value):
                return row
    return None


def find_total_row_after(ws, start_row):
    """Find the next total row after a starting row."""
    if not start_row:
        return None
    for row in range(start_row + 1, ws.max_row + 1):
        row_text = " ".join(
            str(ws.cell(row=row, column=col).value or "")
            for col in range(1, min(20, ws.max_column) + 1)
        )
        if "合计" in row_text:
            return row
    return None


def find_group_child_cols(ws, group_header_row, child_header_row, group_keywords, stop_keywords=None):
    """Find child columns under a grouped header area."""
    stop_keywords = stop_keywords or []
    start_col = None
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=group_header_row, column=col).value
        if value and any(keyword in str(value) for keyword in group_keywords):
            start_col = col
            break
    if not start_col:
        return []

    cols = []
    for col in range(start_col, ws.max_column + 1):
        header = ws.cell(row=child_header_row, column=col).value
        if col > start_col and header and any(keyword in str(header) for keyword in stop_keywords):
            break
        if header not in (None, ""):
            cols.append(col)
    return cols


def clear_constant_cells(ws, row_start, row_end, columns):
    """Clear constants in target columns while keeping formulas."""
    for row in range(row_start, row_end + 1):
        for col in columns:
            cell = ws.cell(row=row, column=col)
            value = cell.value
            if value is not None and not (isinstance(value, str) and value.startswith("=")):
                set_cell_value(ws, row, col, None)


def clear_borders(ws, row_start, row_end, col_start, col_end):
    """Remove dense copied borders from a body area."""
    no_border = Border()
    for row in range(row_start, row_end + 1):
        for col in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=col)
            if not isinstance(cell, MergedCell):
                cell.border = no_border


def highlight_rows(ws, rows, col_start=1, col_end=None):
    """Highlight rows that need manual refresh."""
    fill = PatternFill(fill_type="solid", fgColor="FFFF99")
    col_end = col_end or ws.max_column
    for row in rows:
        if row < 1 or row > ws.max_row:
            continue
        for col in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=col)
            if not isinstance(cell, MergedCell):
                cell.fill = copy(fill)


def clear_blank_borders(ws, row_start=1, row_end=None, col_start=1, col_end=None):
    """Remove borders from blank cells only."""
    no_border = Border()
    row_end = row_end or ws.max_row
    col_end = col_end or ws.max_column
    for row in range(row_start, row_end + 1):
        for col in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=col)
            if isinstance(cell, MergedCell):
                continue
            if cell.value in (None, ""):
                cell.border = no_border


def clear_cell_format(cell, clear_fill=True, clear_border=True):
    """Clear visual-only formatting on a single worksheet cell."""
    if isinstance(cell, MergedCell):
        return
    if clear_border:
        cell.border = Border()
    if clear_fill:
        cell.fill = PatternFill()


def clear_area_format(ws, row_start, row_end, col_start, col_end, clear_fill=True, clear_border=True):
    """Clear fill/border formatting in a bounded area."""
    for row in range(row_start, row_end + 1):
        for col in range(col_start, col_end + 1):
            clear_cell_format(ws.cell(row=row, column=col), clear_fill=clear_fill, clear_border=clear_border)


def row_has_content(ws, row, col_start, col_end):
    """Return True when a row has actual content in a column band."""
    return any(
        ws.cell(row=row, column=col).value not in (None, "")
        for col in range(col_start, col_end + 1)
    )


def update_total_row_formulas(ws, total_row, old_total_row, data_start_row, data_end_row):
    """After inserting detail rows, keep total-row formulas pointed at the new range."""
    if not total_row or total_row == old_total_row:
        return

    def expand_sum(match):
        start_col, start_row, end_col, end_row = match.groups()
        start_row_num = int(start_row)
        end_row_num = int(end_row)
        if start_row_num == data_start_row and end_row_num == old_total_row - 1:
            return f"SUM({start_col}{start_row}:{end_col}{data_end_row})"
        return match.group(0)

    total_ref = re.compile(r"(\$?[A-Z]{1,3}\$?)(%d)(?!\d)" % old_total_row)
    sum_ref = re.compile(r"SUM\((\$?[A-Z]{1,3}\$?)(\d+):(\$?[A-Z]{1,3}\$?)(\d+)\)")

    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=total_row, column=col)
        formula = cell.value
        if not isinstance(formula, str) or not formula.startswith("="):
            continue
        formula = sum_ref.sub(expand_sum, formula)
        formula = total_ref.sub(lambda m: f"{m.group(1)}{total_row}", formula)
        cell.value = formula


def shift_local_formula_refs_after_insert(ws, start_row, offset):
    """Shift simple same-sheet row references in formulas moved by row insertion."""
    if offset <= 0:
        return

    cell_ref = re.compile(r"(?<![A-Za-z0-9_])(\$?[A-Z]{1,3}\$?)(\d+)(?!\d)")

    for row in range(start_row + offset + 1, min(ws.max_row, start_row + offset + 30) + 1):
        for col in range(1, min(ws.max_column, 30) + 1):
            cell = ws.cell(row=row, column=col)
            formula = cell.value
            if not isinstance(formula, str) or not formula.startswith("=") or "!" in formula:
                continue

            def shift_ref(match):
                row_num = int(match.group(2))
                if row_num >= start_row:
                    return f"{match.group(1)}{row_num + offset}"
                return match.group(0)

            cell.value = cell_ref.sub(shift_ref, formula)


def assertion_key(value):
    """Normalize assertion labels used by CRA tables."""
    text = normalize_text(value)
    text = re.split(r"[（(]", text)[0]
    mapping = {
        "完整性": "完整性",
        "存在性": "存在性",
        "计价": "计价",
        "权利和义务": "权利义务",
        "权利义务": "权利义务",
        "列报和披露": "列报",
        "列报": "列报",
    }
    for key, mapped in mapping.items():
        if key in text:
            return mapped
    return text


def numeric_value(value):
    """Return numeric cell values for aggregation; ignore text/formula errors."""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    return None


def process_summary_sheet(ws_new, company_info, bs_date, company_name, warnings_list):
    """Fill common fields on the summary sheet."""
    date_obj = parse_date_value(bs_date)

    fill_adjacent_header_value(ws_new, ["客户名称"], company_name, search_range=(1, 15))

    for row in range(1, min(15, ws_new.max_row) + 1):
        for col in range(1, min(12, ws_new.max_column) + 1):
            value = ws_new.cell(row=row, column=col).value
            if value and "期末" in str(value):
                target = ws_new.cell(row=row, column=col + 1)
                if not isinstance(target, MergedCell):
                    target.value = date_obj if date_obj else bs_date
                    if date_obj:
                        target.number_format = "yyyy/mm/dd"
                break

    fill_adjacent_header_value(ws_new, ["记账本位币", "本位币"], company_info.get("functional_currency"), search_range=(1, 15))
    fill_adjacent_header_value(ws_new, ["适用会计准则", "会计准则"], company_info.get("accounting_standard"), search_range=(1, 15))

    pm_fields = [
        ("PM", ["重要性水平", "PM"]),
        ("TE", ["可容忍误差", "TE"]),
        ("SAD", ["名义金额", "SAD"]),
    ]
    for field, keywords in pm_fields:
        if company_info.get(field):
            fill_adjacent_header_value(ws_new, keywords, company_info[field], search_range=(1, 15))
        else:
            warnings_list.append(f"PMTE信息表中未找到{field}数据，请手动填写")

    cra_data = company_info.get("cra_data", {})
    if cra_data:
        for row in range(1, min(30, ws_new.max_row) + 1):
            assertion = assertion_key(ws_new.cell(row=row, column=3).value)
            if assertion and assertion in cra_data:
                level = cra_data[assertion].get("level")
                if level:
                    set_cell_value(ws_new, row, 4, level)

    for row in range(1, min(30, ws_new.max_row) + 1):
        for col in range(1, min(15, ws_new.max_column) + 1):
            value = ws_new.cell(row=row, column=col).value
            if isinstance(value, (datetime.datetime, datetime.date)):
                ws_new.cell(row=row, column=col).number_format = "yyyy/mm/dd"


def fill_basic_lead_header(ws_new, company_info, bs_date, company_name):
    """Fill common Lead Sheet header fields when no prior Lead Sheet exists."""
    date_obj = parse_date_value(bs_date)

    fill_adjacent_header_value(ws_new, ["客户"], company_name)

    for row in [2, 3, 4]:
        for col in [2, 3]:
            cell_val = ws_new.cell(row=row, column=col).value
            if cell_val and any(x in str(cell_val) for x in ["期末", "资产负债表日"]):
                target = ws_new.cell(row=row, column=col + 1)
                target.value = date_obj if date_obj else bs_date
                if date_obj:
                    target.number_format = "yyyy/mm/dd"
                break
        else:
            continue
        break

    for row in [2, 3, 4, 5]:
        for col in [2, 3]:
            cell_val = ws_new.cell(row=row, column=col).value
            if cell_val and "分析日期" in str(cell_val):
                target = ws_new.cell(row=row, column=col + 1)
                target.value = datetime.datetime.now()
                target.number_format = "yyyy/mm/dd"
                break
        else:
            continue
        break

    fill_adjacent_header_value(ws_new, ["记账本位币", "本位币"], company_info.get("functional_currency"))
    fill_adjacent_header_value(ws_new, ["适用会计准则", "会计准则"], company_info.get("accounting_standard"))
    normalize_lead_date_formats(ws_new)


def extract_company_info_from_pmte(pmte_path, company_name):
    """从PMTE信息表中提取公司信息"""
    if not os.path.exists(pmte_path):
        return {}

    wb = openpyxl.load_workbook(pmte_path, data_only=True)
    info = {
        "PM": None,
        "TE": None,
        "SAD": None,
        "RP": None,
        "Level": None,
    }

    # 读取PMTE Sheet
    if "PMTE" in wb.sheetnames:
        ws = wb["PMTE"]
        for row in range(2, ws.max_row + 1):
            cell_company = ws.cell(row=row, column=1).value
            company_key = normalize_text(company_name)
            cell_company_key = normalize_text(cell_company)
            if cell_company and (
                company_key in cell_company_key
                or cell_company_key in company_key
            ):
                info["Level"] = ws.cell(row=row, column=2).value
                info["RP"] = ws.cell(row=row, column=3).value
                info["PM"] = ws.cell(row=row, column=4).value
                info["TE"] = ws.cell(row=row, column=5).value
                info["SAD"] = ws.cell(row=row, column=6).value
                break

    wb.close()
    return info


def load_cra_data(cra_path, subject_code):
    """
    从CRA等级表中读取指定科目的认定等级和比例

    Args:
        cra_path: CRA等级表路径（与PMTE信息表同一个文件）
        subject_code: 科目代码（如"C"）

    Returns:
        dict: {
            认定名称: {
                "ratio": 比例值（如1.0表示100%）,
                "level": 等级文字（如"Low"）,
                "applicable": 是否适用（True/False）
            },
            ...
        }
    """
    if not os.path.exists(cra_path):
        return {}

    try:
        wb = openpyxl.load_workbook(cra_path, data_only=True)
        if "CRA" not in wb.sheetnames:
            wb.close()
            return {}

        ws = wb["CRA"]
        cra_data = {}

        # 中文等级到英文等级的映射
        level_map = {
            "极低": "Minimal",
            "低": "Low",
            "中等": "Moderate",
            "中": "Moderate",
            "高": "High",
            "很高": "High",
            "极高": "High",
            "不适用": "N/A",
            "N/A": "N/A",
        }

        # 遍历所有行，查找包含 "subject_code." 的认定行
        for row in range(1, ws.max_row + 1):
            c2_val = ws.cell(row=row, column=2).value
            if not c2_val or not isinstance(c2_val, str):
                continue

            # 检查是否以 "subject_code." 开头
            prefix = f"{subject_code}."
            if c2_val.startswith(prefix):
                # 提取认定名称
                # 格式如 "C. 货币资金-存在性"
                parts = c2_val.split("-")
                if len(parts) >= 2:
                    assertion = parts[-1].strip()

                    # 标准化认定名称（去除括号内容）
                    assertion_clean = assertion.split("(")[0].strip()

                    # 统一认定名称映射
                    assertion_map = {
                        "存在性": "存在性",
                        "完整性": "完整性",
                        "计价": "计价",
                        "计价/分摊": "计价",
                        "权利义务": "权利义务",
                        "列报": "列报",
                    }
                    for key, mapped in assertion_map.items():
                        if key in assertion_clean:
                            assertion_clean = mapped
                            break

                    # 读取C3列（是否适用）
                    c3_val = ws.cell(row=row, column=3).value
                    applicable = str(c3_val).strip().upper() != "X" if c3_val is not None else True

                    # 读取C6列（风险等级）
                    c6_val = ws.cell(row=row, column=6).value
                    level_cn = str(c6_val).strip() if c6_val else None
                    level_en = None
                    if level_cn:
                        for cn, en in level_map.items():
                            if cn in level_cn:
                                level_en = en
                                break

                    # 读取C9列（比例值）
                    c9_val = ws.cell(row=row, column=9).value
                    ratio = None
                    if c9_val is not None:
                        try:
                            ratio = float(c9_val)
                        except (ValueError, TypeError):
                            pass

                    cra_data[assertion_clean] = {
                        "ratio": ratio,
                        "level": level_en,
                        "applicable": applicable
                    }

        wb.close()
        return cra_data
    except Exception:
        return {}


def process_lead_sheet(ws_prior, ws_new, ws_prior_values, company_info, bs_date, company_name, lead_config, warnings_list):
    """
    处理Lead Sheet：填写表头 + Roll Forward期初数

    Args:
        ws_prior: 上年底稿的Lead Sheet（含公式）
        ws_new: 新底稿的Lead Sheet
        ws_prior_values: 上年底稿的Lead Sheet（含计算值，data_only=True）
        company_info: 公司信息字典
        bs_date: 资产负债表日期
        company_name: 公司名称
        lead_config: Lead Sheet配置
        warnings_list: 警告列表

    Returns:
        复制的期初数数量
    """
    date_obj = parse_date_value(bs_date)

    # 1. 填写表头（尝试多个位置）
    # 客户名称通常在第2行第2列或附近
    for row in [2, 3, 4]:
        for col in [2, 3, 4]:
            cell_val = ws_new.cell(row=row, column=col).value
            if cell_val and "客户" in str(cell_val):
                ws_new.cell(row=row, column=col + 1, value=company_name)
                break
        else:
            continue
        break

    # 资产负债表日期 - 统一格式
    for row in [2, 3, 4]:
        for col in [2, 3]:
            cell_val = ws_new.cell(row=row, column=col).value
            if cell_val and any(x in str(cell_val) for x in ["期末", "资产负债表日"]):
                # 找到日期单元格（通常在旁边一列）
                target = ws_new.cell(row=row, column=col + 1)
                if target.value is None or "202" in str(target.value):
                    target.value = date_obj if date_obj else bs_date
                    if date_obj:
                        target.number_format = "yyyy/mm/dd"
                break
        else:
            continue
        break

    # 分析日期
    for row in [2, 3, 4, 5]:
        for col in [2, 3]:
            cell_val = ws_new.cell(row=row, column=col).value
            if cell_val and "分析日期" in str(cell_val):
                target = ws_new.cell(row=row, column=col + 1)
                target.value = datetime.datetime.now()
                target.number_format = "yyyy/mm/dd"
                break
        else:
            continue
        break

    # 用户手工输入的基础信息
    fill_adjacent_header_value(
        ws_new,
        ["记账本位币", "本位币"],
        company_info.get("functional_currency")
    )
    fill_adjacent_header_value(
        ws_new,
        ["适用会计准则", "会计准则"],
        company_info.get("accounting_standard")
    )

    normalize_lead_date_formats(ws_new)

    # TE和SAD - 检查PMTE中是否找到
    te_found = False
    sad_found = False

    if company_info.get("TE"):
        for row in [2, 3, 4, 5, 6, 7]:
            for col in [2, 3]:
                cell_val = ws_new.cell(row=row, column=col).value
                if cell_val and ("可容忍误差" in str(cell_val) or "TE" in str(cell_val)):
                    ws_new.cell(row=row, column=col + 1, value=company_info["TE"])
                    te_found = True
                    break
            else:
                continue
            break
    else:
        warnings_list.append("PMTE信息表中未找到TE数据，请手动填写")

    if company_info.get("SAD"):
        for row in [2, 3, 4, 5, 6, 7]:
            for col in [2, 3]:
                cell_val = ws_new.cell(row=row, column=col).value
                if cell_val and ("名义金额" in str(cell_val) or "SAD" in str(cell_val)):
                    ws_new.cell(row=row, column=col + 1, value=company_info["SAD"])
                    sad_found = True
                    break
            else:
                continue
            break
    else:
        warnings_list.append("PMTE信息表中未找到SAD数据，请手动填写")

    # CRA信息 - 填写风险等级(Level)和风险系数(RP)
    if company_info.get("Level"):
        # 在CRA区域查找"实质性"等行，填入Level
        for row in range(13, 25):
            c2_val = ws_new.cell(row=row, column=2).value
            if c2_val and "实质性" in str(c2_val):
                # C列填入Level（风险等级，如"Low", "Moderate", "High"）
                ws_new.cell(row=row, column=3).value = company_info["Level"]
                break
    else:
        warnings_list.append("PMTE信息表中未找到Level数据，请手动填写风险等级")

    if company_info.get("RP"):
        # 在CRA区域查找"实质性"等行，填入RP
        for row in range(13, 25):
            c2_val = ws_new.cell(row=row, column=2).value
            if c2_val and "实质性" in str(c2_val):
                # D列填入RP（风险系数百分比）
                ws_new.cell(row=row, column=4).value = company_info["RP"]
                break
    else:
        warnings_list.append("PMTE信息表中未找到RP数据，请手动填写风险系数")

    # 新增：根据CRA等级表动态填写每个认定的等级和比例
    cra_data = company_info.get("cra_data", {})
    if cra_data:
        # 获取C5的值（Threshold，用于计算比例）
        c5_value = ws_new.cell(row=5, column=3).value
        # C5可能是公式，尝试获取数值
        try:
            if c5_value is not None:
                c5_num = float(c5_value)
            else:
                c5_num = None
        except (ValueError, TypeError):
            c5_num = None

        # 遍历CRA区域的认定行（R15-R19）
        for row in range(15, 20):
            c2_val = ws_new.cell(row=row, column=2).value
            if not c2_val:
                continue

            # 提取认定名称（去掉括号内容，同时处理全角和半角括号）
            import re
            assertion_name = re.split(r'[（(]', str(c2_val))[0].strip()
            # 标准化认定名称
            assertion_map = {
                "存在性": "存在性",
                "完整性": "完整性",
                "计价": "计价",
                "计价/分摊": "计价",
                "权利义务": "权利义务",
                "列报": "列报",
            }
            for key, mapped in assertion_map.items():
                if key in assertion_name:
                    assertion_name = mapped
                    break

            # 在CRA数据中查找对应认定
            if assertion_name in cra_data:
                data = cra_data[assertion_name]

                # 填入等级（C列）
                if data.get("level"):
                    ws_new.cell(row=row, column=3).value = data["level"]

                # 如果比例与公式默认不同，覆盖D列
                if data.get("ratio") is not None:
                    # 计算当前等级对应的默认比例
                    current_level = data.get("level", "")
                    default_ratio = None
                    if current_level == "Minimal":
                        default_ratio = 1.0
                    elif current_level == "Low":
                        default_ratio = 0.75
                    elif current_level == "Moderate":
                        default_ratio = 0.50
                    elif current_level == "High":
                        default_ratio = 0.25

                    # 如果用户比例与默认比例不同，覆盖D列
                    if default_ratio is None or abs(data["ratio"] - default_ratio) > 0.001:
                        # 将D列的公式替换为计算值
                        ratio_value = data["ratio"]
                        # 如果C5有值，计算实际金额；否则直接写入比例值
                        if c5_num is not None:
                            d_value = c5_num * ratio_value
                        else:
                            d_value = ratio_value
                        ws_new.cell(row=row, column=4).value = d_value

    # 2. Roll Forward期初数（从计算值工作表读取）
    overwrite_opening_formulas = lead_config.get("overwrite_opening_formulas", False)

    header_search_text = lead_config.get("header_search_text", "期末审定数")
    prior_header_row = find_header_row(ws_prior_values, header_search_text, (1, 80))
    new_header_row = find_header_row(ws_new, header_search_text, (1, 80))

    if not prior_header_row or not new_header_row:
        return 0

    prior_closing_col = (
        find_header_col(ws_prior_values, prior_header_row, ["期末审定数"])
        or lead_config.get("closing_col", 9)
    )
    opening_col = (
        find_header_col(ws_new, new_header_row, ["上期末审定数", "期初审定数", "上年审定数"])
        or lead_config.get("opening_col", 10)
    )

    descriptor_rules = [
        (["账套名称", "账套编码", "账套"], ["账套名称", "账套编码", "账套"]),
        (["总账科目编码", "科目编码"], ["总账科目编码", "科目编码"]),
        (["科目名称", "报表科目", "项目名称"], ["科目名称", "报表科目", "项目名称"]),
        (["索引号", "索引"], ["索引号", "索引"]),
    ]
    descriptor_cols = []
    for prior_keywords, new_keywords in descriptor_rules:
        source_col = find_header_col_near(ws_prior_values, prior_header_row, prior_keywords)
        target_col = find_header_col_near(ws_new, new_header_row, new_keywords)
        if source_col and target_col:
            descriptor_cols.append((source_col, target_col))
    if not descriptor_cols:
        descriptor_cols = [(2, 2), (3, 3), (4, 4)]

    prior_data_rows = []
    for row in range(prior_header_row + 1, ws_prior_values.max_row + 1):
        if row_contains_any(ws_prior_values, row, ["合计"], columns=range(2, min(8, ws_prior_values.max_column) + 1)):
            break
        if row_contains_any(ws_prior_values, row, ["check with", "Diff", "波动说明", "Notes：", "Notes:"], columns=range(2, min(8, ws_prior_values.max_column) + 1)):
            break
        if row_contains_any(ws_prior_values, row, ["Rx"], columns=range(2, min(8, ws_prior_values.max_column) + 1)):
            continue

        closing_val = ws_prior_values.cell(row=row, column=prior_closing_col).value
        descriptor_values = {
            target_col: ws_prior_values.cell(row=row, column=source_col).value
            for source_col, target_col in descriptor_cols
        }
        if closing_val is not None or any(value not in (None, "") for value in descriptor_values.values()):
            prior_data_rows.append((row, descriptor_values, closing_val))

    new_data_start_row = new_header_row + 1
    new_total_row = None
    for row in range(new_data_start_row, ws_new.max_row + 1):
        if row_contains_any(ws_new, row, ["合计"], columns=range(2, min(8, ws_new.max_column) + 1)):
            new_total_row = row
            break

    template_data_count = 0
    if new_total_row:
        for row in range(new_data_start_row, new_total_row):
            if row_contains_any(ws_new, row, ["Rx", "A3", "Diff", "波动说明"], columns=range(2, min(8, ws_new.max_column) + 1)):
                continue
            if any(
                ws_new.cell(row=row, column=col).value not in (None, "")
                for col in sorted({col for _, col in descriptor_cols} | {opening_col})
            ):
                template_data_count += 1
    else:
        template_data_count = len(prior_data_rows)

    new_total_row_adjusted = new_total_row
    rows_to_insert = 0
    if new_total_row and len(prior_data_rows) > template_data_count:
        rows_to_insert = len(prior_data_rows) - template_data_count
        source_row = max(new_data_start_row, new_total_row - 1)
        insert_rows_preserving_sheet_metadata(ws_new, new_total_row, rows_to_insert)
        for offset in range(rows_to_insert):
            copy_row_shape(ws_new, source_row, new_total_row + offset, translate_formula=True)

        new_total_row_adjusted = new_total_row + rows_to_insert
        update_total_row_formulas(
            ws_new,
            new_total_row_adjusted,
            new_total_row,
            new_data_start_row,
            new_total_row_adjusted - 1,
        )
        shift_local_formula_refs_after_insert(ws_new, new_total_row, rows_to_insert)

    copied = 0
    for idx, (prior_row, descriptor_values, closing_val) in enumerate(prior_data_rows):
        target_row = new_data_start_row + idx if new_total_row else prior_row

        for target_col, value in descriptor_values.items():
            set_cell_value(ws_new, target_row, target_col, value)

        existing_value = ws_new.cell(row=target_row, column=opening_col).value
        keeps_formula = isinstance(existing_value, str) and existing_value.startswith("=")
        if closing_val is not None and (overwrite_opening_formulas or not keeps_formula):
            if set_cell_value(ws_new, target_row, opening_col, closing_val):
                copied += 1

    clear_current_cols = lead_config.get("clear_current_period_cols", [])
    if clear_current_cols:
        for idx, (prior_row, descriptor_values, closing_val) in enumerate(prior_data_rows):
            target_row = new_data_start_row + idx if new_total_row else prior_row
            if not any(value not in (None, "") for value in descriptor_values.values()):
                continue
            for col in clear_current_cols:
                set_cell_value(ws_new, target_row, col, None)

    # 3. 处理For Disclosure表（如果存在）
    # 查找"For Disclosure"相关行
    disclosure_start_row = None
    for row in range(1, ws_prior_values.max_row + 1):
        for col in range(1, min(10, ws_prior_values.max_column + 1)):
            val = ws_prior_values.cell(row=row, column=col).value
            if val and "Disclosure" in str(val):
                disclosure_start_row = row
                break
        if disclosure_start_row:
            break

    if disclosure_start_row:
        # 在For Disclosure区域内查找包含"期末审定数"的表头行（旧文件）
        fd_prior_header_row = None
        fd_prior_closing_col = None
        fd_prior_opening_col = None

        for row in range(disclosure_start_row, min(disclosure_start_row + 10, ws_prior_values.max_row + 1)):
            for col in range(1, min(20, ws_prior_values.max_column + 1)):
                val = ws_prior_values.cell(row=row, column=col).value
                if val and "期末审定数" in str(val):
                    fd_prior_header_row = row
                    fd_prior_closing_col = col
                    # 查找"期初审定数"列（通常在下一列）
                    for next_col in range(col + 1, min(col + 5, ws_prior_values.max_column + 1)):
                        next_val = ws_prior_values.cell(row=row, column=next_col).value
                        if next_val and "期初" in str(next_val):
                            fd_prior_opening_col = next_col
                            break
                    break
            if fd_prior_header_row:
                break

        # 在新文件中查找For Disclosure区域的表头行（通过"Disclosure"和"报表科目"定位）
        fd_new_header_row = None
        fd_new_opening_col = None

        for row in range(1, ws_new.max_row + 1):
            val = ws_new.cell(row=row, column=2).value  # C2通常是"报表科目"
            if val and ("报表科目" in str(val) or "项目名称" in str(val)):
                # 检查上方几行是否有"Disclosure"或"For"
                found_disclosure = False
                for check_row in range(max(1, row - 5), row):
                    for check_col in range(1, 5):
                        check_val = ws_new.cell(row=check_row, column=check_col).value
                        if check_val and "Disclosure" in str(check_val):
                            found_disclosure = True
                            break
                    if found_disclosure:
                        break
                if found_disclosure:
                    fd_new_header_row = row
                    # 查找"期初"列（在表头行中搜索"期初"）
                    # 注意：新文件可能包含公式而非文本，所以尝试多种方式
                    for col in range(1, min(20, ws_new.max_column + 1)):
                        col_val = ws_new.cell(row=row, column=col).value
                        if col_val and "期初" in str(col_val):
                            fd_new_opening_col = col
                            break
                    # 如果没找到"期初"，尝试找"期末审定数"然后+1
                    if not fd_new_opening_col:
                        for col in range(1, min(20, ws_new.max_column + 1)):
                            col_val = ws_new.cell(row=row, column=col).value
                            if col_val and "期末审定数" in str(col_val):
                                fd_new_opening_col = col + 1
                                break
                    # 如果还是没找到，默认C4（通常For Disclosure表结构固定）
                    if not fd_new_opening_col:
                        fd_new_opening_col = 4
                    break

        # 按项目名称匹配复制
        if fd_prior_header_row and fd_new_header_row and fd_prior_closing_col and fd_new_opening_col:
            # 构建旧文件的报表科目→期末审定数字典
            prior_items = {}
            for row in range(fd_prior_header_row + 1, ws_prior_values.max_row + 1):
                item_name = ws_prior_values.cell(row=row, column=2).value  # C2通常是报表科目
                if item_name and str(item_name).strip():
                    val = ws_prior_values.cell(row=row, column=fd_prior_closing_col).value
                    if val is not None:
                        try:
                            prior_items[str(item_name).strip()] = float(val)
                        except (ValueError, TypeError):
                            prior_items[str(item_name).strip()] = val

            # 在新文件中查找匹配的项目并填充
            for row in range(fd_new_header_row + 1, ws_new.max_row + 1):
                item_name = ws_new.cell(row=row, column=2).value
                if item_name and str(item_name).strip() in prior_items:
                    ws_new.cell(row=row, column=fd_new_opening_col).value = prior_items[str(item_name).strip()]
                    copied += 1

    return copied


def process_k01(ws_prior, ws_new, k01_config):
    """
    处理K.01 Agree SL to GL：复制表头 + 年初余额

    Args:
        ws_prior: 上年底稿的K.01工作表（含计算值）
        ws_new: 新底稿的K.01工作表
        k01_config: K.01配置

    Returns:
        复制的数据数量
    """
    if not k01_config.get("has_k01", False):
        return 0

    # 1. 复制表头行
    header_row = k01_config.get("header_row", 10)
    for col in range(1, ws_prior.max_column + 1):
        prior_cell = ws_prior.cell(row=header_row, column=col)
        new_cell = ws_new.cell(row=header_row, column=col)
        if isinstance(new_cell, MergedCell):
            continue
        if prior_cell.value is not None:
            new_cell.value = prior_cell.value

    roll_forward_groups = k01_config.get("roll_forward_groups", [])
    if roll_forward_groups:
        copied = 0
        for group in roll_forward_groups:
            group_name = group.get("group")
            source_detail = group.get("source_detail", "期末余额")
            target_detail = group.get("target_detail", "期初余额")
            value_cols = group.get("value_cols", [5, 7, 9])

            source_row = find_group_detail_row(ws_prior, group_name, source_detail)
            target_row = find_group_detail_row(ws_new, group_name, target_detail)
            if not source_row or not target_row:
                continue

            for col in value_cols:
                value = ws_prior.cell(row=source_row, column=col).value
                if value is not None:
                    ws_new.cell(row=target_row, column=col).value = value
                    copied += 1
        return copied

    # 2. 复制年初余额数据
    opening_balance_rows = k01_config.get("opening_balance_rows", [])
    copied = 0

    for balance_row in opening_balance_rows:
        for col in range(1, ws_prior.max_column + 1):
            prior_cell = ws_prior.cell(row=balance_row, column=col)
            new_cell = ws_new.cell(row=balance_row, column=col)
            if prior_cell.value is not None:
                new_cell.value = prior_cell.value
                copied += 1

    return copied


def find_group_detail_row(ws, group_name, detail_name):
    """Find a row by group label in column B and detail label in column C."""
    group_key = normalize_text(group_name)
    detail_key = normalize_text(detail_name)
    if not group_key or not detail_key:
        return None

    in_group = False
    for row in range(1, ws.max_row + 1):
        group_value = normalize_text(ws.cell(row=row, column=2).value)
        detail_value = normalize_text(ws.cell(row=row, column=3).value)

        if group_value:
            in_group = group_key in group_value or group_value in group_key

        if in_group and detail_key in detail_value:
            return row

    return None


def find_schedule_section_row(ws, section_name, detail_name):
    """Find a row in a roll-forward schedule by section in column B and detail in column C."""
    section_key = normalize_text(section_name)
    detail_key = normalize_text(detail_name)
    in_section = False

    for row in range(1, ws.max_row + 1):
        section_value = normalize_text(ws.cell(row=row, column=2).value)
        detail_value = normalize_text(ws.cell(row=row, column=3).value)

        if section_value:
            in_section = section_key in section_value or section_value in section_key

        if in_section and detail_key in detail_value:
            return row

    return None


def find_schedule_group_columns(ws):
    """Return roll-forward schedule category columns keyed by category label."""
    groups = []
    for col in range(1, ws.max_column + 1):
        label = ws.cell(row=2, column=col).value
        if not label:
            continue
        label_text = str(label).strip()
        if not label_text:
            continue
        book_header = normalize_text(ws.cell(row=3, column=col - 1).value) if col > 1 else ""
        audit_header = normalize_text(ws.cell(row=3, column=col + 1).value)
        if "账面数" not in book_header or "审定数" not in audit_header:
            continue

        groups.append({
            "name": label_text,
            "name_key": normalize_text(label_text),
            "book_col": col - 1,
            "adjust_col": col,
            "audit_col": col + 1,
        })

    return groups


def get_schedule_total_col(ws):
    for group in find_schedule_group_columns(ws):
        if "合计" in group["name_key"]:
            return group["audit_col"]
    return None


def l1_category_key(name):
    key = normalize_text(name)
    if "土地" in key:
        return "土地"
    if "非专利" in key:
        return "非专利"
    if "专利" in key:
        return "专利"
    if "软件" in key or "计算机" in key:
        return "软件"
    if "其他" in key:
        return "其他"
    return key


def process_l1_from_rollforward_schedule(ws_schedule, ws_new_lead, ws_new_k01, subject_config):
    """Populate L1 from an LAR/LRA roll-forward schedule."""
    copied = 0

    total_audit_col = get_schedule_total_col(ws_schedule)
    sections = [
        (["原值"], "无形资产", "1701"),
        (["累计摊销", "累计折旧"], "累计摊销", "1702"),
        (["减值准备"], "减值准备", "1703"),
    ]

    # L1.00 Lead sheet: PY values and GL account codes.
    lead_rows = {}
    for row in range(1, ws_new_lead.max_row + 1):
        label = normalize_text(ws_new_lead.cell(row=row, column=3).value)
        for section_names, lead_label, account_code in sections:
            if normalize_text(lead_label) in label:
                lead_rows[lead_label] = row

    if total_audit_col:
        for section_names, lead_label, account_code in sections:
            source_row = None
            for section_name in section_names:
                source_row = find_schedule_section_row(ws_schedule, section_name, "年末余额")
                if source_row:
                    break
            target_row = lead_rows.get(section_name)
            target_row = lead_rows.get(lead_label)
            if source_row and target_row:
                set_cell_value(ws_new_lead, target_row, 2, account_code)
                value = ws_schedule.cell(row=source_row, column=total_audit_col).value
                if value is not None:
                    set_cell_value(ws_new_lead, target_row, 10, value)
                    copied += 1

    # L1.01.1 Agree SL to GL: prior year-end rolls to current opening rows.
    schedule_groups = [g for g in find_schedule_group_columns(ws_schedule) if "合计" not in g["name_key"]]
    target_order = ["土地", "非专利", "专利", "软件", "其他"]
    target_labels = {
        "土地": "土地使用权",
        "非专利": "非专利技术",
        "专利": "专利权",
        "软件": "软件",
        "其他": "其他",
    }
    target_groups = {}
    for key, col in zip(target_order, [6, 9, 12, 15, 18]):
        target_groups[key] = {
            "name": target_labels[key],
            "name_key": key,
            "book_col": col - 1,
            "adjust_col": col,
        }

    target_values = {}

    for schedule_group in schedule_groups:
        target = target_groups.get(l1_category_key(schedule_group["name"]))
        if not target:
            continue

        target_key = target["name_key"]
        if target_key not in target_values:
            target_values[target_key] = {
                "target": target,
                "source_names": [],
                "values": {},
            }
        target_values[target_key]["source_names"].append(schedule_group["name"])

        for section_names, target_row in [(["原值"], 12), (["累计摊销", "累计折旧"], 18), (["减值准备"], 23)]:
            source_row = None
            for section_name in section_names:
                source_row = find_schedule_section_row(ws_schedule, section_name, "年末余额")
                if source_row:
                    break
            if not source_row:
                continue

            values = target_values[target_key]["values"]
            values.setdefault(target_row, {"book": 0, "adjust": 0, "book_seen": False, "adjust_seen": False})
            book_value = ws_schedule.cell(row=source_row, column=schedule_group["book_col"]).value
            adjust_value = ws_schedule.cell(row=source_row, column=schedule_group["adjust_col"]).value
            book_value = numeric_value(book_value)
            adjust_value = numeric_value(adjust_value)
            if book_value is not None:
                values[target_row]["book"] += book_value
                values[target_row]["book_seen"] = True
                copied += 1
            if adjust_value is not None:
                values[target_row]["adjust"] += adjust_value
                values[target_row]["adjust_seen"] = True
                copied += 1

    for item in target_values.values():
        target = item["target"]
        source_names = list(dict.fromkeys(item["source_names"]))
        ws_new_k01.cell(row=10, column=target["adjust_col"]).value = (
            source_names[0] if len(source_names) == 1 else "+".join(source_names)
        )
        for target_row, values in item["values"].items():
            if values["book_seen"]:
                ws_new_k01.cell(row=target_row, column=target["book_col"]).value = values["book"]
            if values["adjust_seen"]:
                ws_new_k01.cell(row=target_row, column=target["adjust_col"]).value = values["adjust"]

    return copied


def process_l103_policy_table(ws_prior, ws_new):
    """Roll L1.03 table 2 prior-year policy information into the new workbook."""
    copied = 0
    prior_rows = {}

    for row in range(1, ws_prior.max_row + 1):
        category = ws_prior.cell(row=row, column=2).value
        if not category:
            continue
        category_key = normalize_text(category)
        if not category_key or "资产类别" in category_key or "表2" in category_key:
            continue
        current_life = ws_prior.cell(row=row, column=3).value
        reason = ws_prior.cell(row=row, column=7).value
        if current_life is not None or reason is not None:
            prior_rows[category_key] = {
                "life": current_life,
                "reason": reason,
            }

    for row in range(1, ws_new.max_row + 1):
        category = ws_new.cell(row=row, column=2).value
        if not category:
            continue
        category_key = normalize_text(category)
        if category_key not in prior_rows:
            continue

        data = prior_rows[category_key]
        if data["life"] is not None:
            set_cell_value(ws_new, row, 4, data["life"])
            copied += 1
        if data["reason"] is not None:
            set_cell_value(ws_new, row, 7, data["reason"])
            copied += 1

    return copied


def process_n_lead_turnover_analysis(ws_prior_values, ws_new):
    """Roll N.00 table 2 prior-year turnover analysis values into PY column."""
    section_text = "表2 应付账款周转率分析"
    prior_section_row = find_row_containing(ws_prior_values, section_text, (1, ws_prior_values.max_row))
    new_section_row = find_row_containing(ws_new, section_text, (1, ws_new.max_row))
    if not prior_section_row or not new_section_row:
        return 0

    prior_source_col = 4
    new_target_col = None
    for row in range(new_section_row, min(new_section_row + 5, ws_new.max_row) + 1):
        for col in range(1, min(12, ws_new.max_column) + 1):
            value = ws_new.cell(row=row, column=col).value
            if value and "PY" in str(value):
                new_target_col = col
                break
        if new_target_col:
            break
    if not new_target_col:
        new_target_col = 5

    prior_values = {}
    for row in range(prior_section_row + 1, min(prior_section_row + 20, ws_prior_values.max_row) + 1):
        label = normalize_text(ws_prior_values.cell(row=row, column=2).value)
        if not label:
            continue
        value = ws_prior_values.cell(row=row, column=prior_source_col).value
        if value is not None:
            prior_values[label] = value

    copied = 0
    for row in range(new_section_row + 1, min(new_section_row + 20, ws_new.max_row) + 1):
        label = normalize_text(ws_new.cell(row=row, column=2).value)
        if label in prior_values:
            set_cell_value(ws_new, row, new_target_col, prior_values[label])
            copied += 1

    return copied


def process_n_detail_sheet(ws_prior_formula, ws_prior_values, ws_new, bs_date, ws_prior_lead=None, ws_new_lead=None):
    """Copy N.01.01 detail sheet and roll prior closing values into the PY column."""
    clone_worksheet_contents(ws_prior_formula, ws_new)

    date_obj = parse_date_value(bs_date)
    if date_obj:
        set_cell_value(ws_new, 5, 8, date_obj)
        ws_new.cell(row=5, column=8).number_format = "yyyy/mm/dd"
        prior_date = datetime.datetime(date_obj.year - 1, 12, 31)
        set_cell_value(ws_new, 5, 14, prior_date)
        ws_new.cell(row=5, column=14).number_format = "yyyy/mm/dd"
        set_cell_value(ws_new, 225, 3, date_obj)
        ws_new.cell(row=225, column=3).number_format = "yyyy/mm/dd"
        set_cell_value(ws_new, 225, 5, prior_date)
        ws_new.cell(row=225, column=5).number_format = "yyyy/mm/dd"
        set_cell_value(ws_new, 257, 4, date_obj)
        ws_new.cell(row=257, column=4).number_format = "yyyy/mm/dd"

    header_row = find_header_row(ws_new, "期末审定数", (1, 30))
    if not header_row:
        return 0

    total_row = find_total_row_after(ws_new, header_row)
    data_end_row = (total_row - 1) if total_row else ws_new.max_row
    closing_col = find_header_col(ws_new, header_row, ["期末审定数"]) or 13
    py_col = find_header_col(ws_new, header_row, ["上期末审定数", "上年数"]) or 14
    current_cols = [
        col for col in [
            find_header_col(ws_new, header_row, ["原币金额"]),
            find_header_col(ws_new, header_row, ["期末账面数"]),
            find_header_col(ws_new, header_row, ["本期审计调整编号"]),
            find_header_col(ws_new, header_row, ["审计调整"]),
            find_header_col(ws_new, header_row, ["重分类调整"]),
            closing_col,
        ]
        if col
    ]
    aging_cols = find_group_child_cols(
        ws_new,
        header_row - 1,
        header_row,
        ["账龄"],
        stop_keywords=["check", "检查"],
    )
    current_cols = sorted(set(current_cols + aging_cols))

    copied = 0
    for row in range(header_row + 1, data_end_row + 1):
        source_value = ws_prior_values.cell(row=row, column=closing_col).value
        if source_value is not None:
            set_cell_value(ws_new, row, py_col, source_value)
            copied += 1

    clear_constant_cells(ws_new, header_row + 1, data_end_row, current_cols)

    if ws_prior_lead is not None and ws_new_lead is not None:
        prior_header_row = find_header_row(ws_prior_lead, "期末审定数", (1, 100))
        new_header_row = find_header_row(ws_new_lead, "期末审定数", (1, 120))
        prior_total_row = find_total_row_after(ws_prior_lead, prior_header_row)
        new_total_row = find_total_row_after(ws_new_lead, new_header_row)
        if prior_total_row and new_total_row and prior_total_row != new_total_row:
            for row in range(1, ws_new.max_row + 1):
                for col in range(1, ws_new.max_column + 1):
                    cell = ws_new.cell(row=row, column=col)
                    formula = cell.value
                    if isinstance(formula, str) and formula.startswith("=") and "'N.00 Lead sheet'!" in formula:
                        cell.value = re.sub(
                            rf"('N\.00 Lead sheet'!\$?[A-Z]{{1,3}})\$?{prior_total_row}(?!\d)",
                            rf"\g<1>{new_total_row}",
                            formula,
                        )

    wording_start_row = None
    for row in range(1, ws_new.max_row + 1):
        row_text = " ".join(
            str(ws_new.cell(row=row, column=col).value or "")
            for col in range(1, min(ws_new.max_column, 10) + 1)
        )
        if "对于单项变动金额" in row_text:
            wording_start_row = row
            break
    if wording_start_row:
        wording_rows = []
        for row in range(wording_start_row, ws_new.max_row + 1):
            if any(
                ws_new.cell(row=row, column=col).value not in (None, "")
                for col in range(2, min(ws_new.max_column, 8) + 1)
            ):
                wording_rows.append(row)

        clear_area_format(
            ws_new,
            wording_start_row,
            ws_new.max_row,
            9,
            min(ws_new.max_column, 24),
            clear_fill=True,
            clear_border=True,
        )
        for row in range(wording_start_row, ws_new.max_row + 1):
            if not row_has_content(ws_new, row, 2, min(ws_new.max_column, 8)):
                clear_area_format(
                    ws_new,
                    row,
                    row,
                    1,
                    min(ws_new.max_column, 24),
                    clear_fill=True,
                    clear_border=True,
                )

        highlight_rows(ws_new, wording_rows, 2, min(ws_new.max_column, 8))

    ws_new.sheet_view.showGridLines = False

    return copied


def generate_output_filename(template_name, bs_date, company_name):
    """生成输出文件名"""
    output_name = template_name
    # 替换各种日期占位符
    output_name = re.sub(r'202[A-Za-z]{4,6}', bs_date.replace('-', ''), output_name)
    output_name = re.sub(r'\d{4}-\d{2}-\d{2}', bs_date, output_name)
    output_name = re.sub(r'\d{8}', bs_date.replace('-', ''), output_name)

    # 替换公司名称占位符
    output_name = re.sub(r'XYZ公司', company_name, output_name)
    output_name = re.sub(r'[<>:"/\\|?*]', '_', output_name)

    return output_name


def process_single_subject(subject_code, template_path, prior_path, pmte_path,
                           company_name, bs_date, output_dir, subject_config,
                           cra_path=None, functional_currency=None, accounting_standard=None):
    """
    处理单个科目的Roll Forward

    Args:
        subject_code: 科目代码（如"K1"）
        template_path: 标准模板路径
        prior_path: 上年底稿路径
        pmte_path: PMTE信息表路径
        company_name: 公司名称
        bs_date: 资产负债表日期（格式: YYYY-MM-DD）
        output_dir: 输出目录
        subject_config: 科目配置
        cra_path: CRA等级表路径（可选，默认与pmte_path相同）

    Returns:
        (success: bool, message: str, output_path: str, warnings: list)
    """
    warnings_list = []

    try:
        # 1. 提取公司信息
        company_info = extract_company_info_from_pmte(pmte_path, company_name)
        company_info["functional_currency"] = functional_currency
        company_info["accounting_standard"] = accounting_standard

        # 1.1 加载CRA等级表数据
        # 如果未指定CRA路径，尝试从PMTE同目录查找CRA文件
        cra_data = None
        if cra_path and os.path.exists(cra_path):
            cra_data = load_cra_data(cra_path, subject_code)
        elif pmte_path and os.path.exists(pmte_path):
            cra_data = load_cra_data(pmte_path, subject_code)

        if cra_data:
            company_info["cra_data"] = cra_data

        # 2. 生成输出文件名
        template_name = os.path.basename(template_path)
        output_name = generate_output_filename(template_name, bs_date, company_name)
        output_path = os.path.join(output_dir, output_name)

        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)

        # 3. 复制标准模板
        shutil.copy2(template_path, output_path)
        # 移除只读属性（Windows）
        os.chmod(output_path, 0o666)

        # 4. 打开新底稿和上年底稿
        # wb_prior_formula: 含公式，用于复制表头结构
        # wb_prior_values: 含计算值，用于获取期末审定数
        wb_new = openpyxl.load_workbook(output_path)
        wb_prior_formula = openpyxl.load_workbook(prior_path, data_only=False)
        wb_prior_values = openpyxl.load_workbook(prior_path, data_only=True)

        try:
            lead_config = subject_config.get("lead_sheet", {})
            k01_config = subject_config.get("k01", {})

            if "汇总" in wb_new.sheetnames:
                process_summary_sheet(wb_new["汇总"], company_info, bs_date, company_name, warnings_list)

            # 5. 处理Lead Sheet
            lead_sheet_name = lead_config.get("sheet_name", "")
            if lead_sheet_name and lead_sheet_name in wb_new.sheetnames and lead_sheet_name in wb_prior_formula.sheetnames:
                ws_new_lead = wb_new[lead_sheet_name]
                ws_prior_formula_lead = wb_prior_formula[lead_sheet_name]
                ws_prior_values_lead = wb_prior_values[lead_sheet_name]
                process_lead_sheet(ws_prior_formula_lead, ws_new_lead, ws_prior_values_lead,
                                   company_info, bs_date, company_name, lead_config, warnings_list)
            elif lead_sheet_name and lead_sheet_name in wb_new.sheetnames:
                ws_new_lead = wb_new[lead_sheet_name]
                fill_basic_lead_header(ws_new_lead, company_info, bs_date, company_name)

            # 6. 处理K.01
            if k01_config.get("has_k01", False):
                k01_sheet_name = k01_config.get("sheet_name", "")
                if k01_sheet_name and k01_sheet_name in wb_new.sheetnames and k01_sheet_name in wb_prior_values.sheetnames:
                    ws_new_k01 = wb_new[k01_sheet_name]
                    ws_prior_k01 = wb_prior_values[k01_sheet_name]
                    process_k01(ws_prior_k01, ws_new_k01, k01_config)

            # 6.1 处理非标准L1 LAR/LRA后推明细表
            if subject_code == "L1" and "后推明细表" in wb_prior_values.sheetnames:
                lead_sheet_name = lead_config.get("sheet_name", "")
                k01_sheet_name = k01_config.get("sheet_name", "")
                if lead_sheet_name in wb_new.sheetnames and k01_sheet_name in wb_new.sheetnames:
                    process_l1_from_rollforward_schedule(
                        wb_prior_values["后推明细表"],
                        wb_new[lead_sheet_name],
                        wb_new[k01_sheet_name],
                        subject_config
                    )

            if subject_code == "L1":
                prior_l103_sheet = next((s for s in wb_prior_formula.sheetnames if "L1.03" in s), None)
                new_l103_sheet = next((s for s in wb_new.sheetnames if "L1.03" in s), None)
                if prior_l103_sheet and new_l103_sheet:
                    process_l103_policy_table(
                        wb_prior_formula[prior_l103_sheet],
                        wb_new[new_l103_sheet]
                    )

            if subject_code == "N":
                lead_sheet_name = lead_config.get("sheet_name", "")
                if lead_sheet_name in wb_new.sheetnames and lead_sheet_name in wb_prior_values.sheetnames:
                    process_n_lead_turnover_analysis(
                        wb_prior_values[lead_sheet_name],
                        wb_new[lead_sheet_name]
                    )

                detail_sheet_name = "N.01.01明细账"
                if (
                    detail_sheet_name in wb_new.sheetnames
                    and detail_sheet_name in wb_prior_formula.sheetnames
                    and detail_sheet_name in wb_prior_values.sheetnames
                ):
                    ws_prior_lead = wb_prior_formula[lead_sheet_name] if lead_sheet_name in wb_prior_formula.sheetnames else None
                    ws_new_lead = wb_new[lead_sheet_name] if lead_sheet_name in wb_new.sheetnames else None
                    process_n_detail_sheet(
                        wb_prior_formula[detail_sheet_name],
                        wb_prior_values[detail_sheet_name],
                        wb_new[detail_sheet_name],
                        bs_date,
                        ws_prior_lead,
                        ws_new_lead
                    )

            # 7. 处理子表（如U_exp的财务费用子表）
            sub_sheets = subject_config.get("sub_sheets", [])
            for sub_sheet in sub_sheets:
                sub_sheet_name = sub_sheet.get("sheet_name", "")
                if sub_sheet_name and sub_sheet_name in wb_new.sheetnames and sub_sheet_name in wb_prior_values.sheetnames:
                    ws_new_sub = wb_new[sub_sheet_name]
                    ws_prior_sub = wb_prior_values[sub_sheet_name]

                    # 查找表头行
                    header_search_text = sub_sheet.get("header_search_text", "期末审定数")
                    header_row = find_header_row(ws_prior_sub, header_search_text, (1, 80))

                    if header_row:
                        closing_col = sub_sheet.get("closing_col", 9)
                        opening_col = sub_sheet.get("opening_col", 10)

                        for row in range(header_row + 1, ws_prior_sub.max_row + 1):
                            prior_cell = ws_prior_sub.cell(row=row, column=closing_col)
                            new_cell = ws_new_sub.cell(row=row, column=opening_col)

                            if prior_cell.value is not None:
                                new_cell.value = prior_cell.value

            # 8. 保存
            wb_new.calculation.fullCalcOnLoad = True
            wb_new.calculation.forceFullCalc = True
            wb_new.calculation.calcMode = "auto"
            wb_new.save(output_path)

            # 生成警告消息
            warning_msg = ""
            if warnings_list:
                warnings_list[:] = list(dict.fromkeys(warnings_list))
                warning_msg = "; ".join(warnings_list)

            return True, f"处理成功{(' - ' + warning_msg if warning_msg else '')}", output_path, warnings_list

        finally:
            wb_new.close()
            wb_prior_formula.close()
            wb_prior_values.close()

    except Exception as e:
        return False, f"处理失败: {str(e)}", None, warnings_list


def process_multiple_subjects(subject_codes, template_dir, prior_dir, pmte_path,
                              company_name, bs_date, output_dir, config_path=None,
                              functional_currency=None, accounting_standard=None):
    """
    批量处理多个科目
    """
    config_manager = SubjectConfig(config_path)
    results = []

    # 计算上年日期
    prior_year = str(int(bs_date[:4]) - 1)

    for subject_code in subject_codes:
        subject_config = config_manager.get_subject(subject_code)
        if not subject_config:
            results.append((subject_code, False, "找不到科目配置", None, []))
            continue

        # 查找标准模板
        template_file = subject_config.get("template_file", "")
        template_path = os.path.join(template_dir, template_file)
        if not os.path.exists(template_path):
            results.append((subject_code, False, f"找不到标准模板: {template_file}", None, []))
            continue

        # 查找上年底稿
        prior_path = find_prior_file(prior_dir, subject_code, prior_year, subject_config)
        if not prior_path:
            results.append((subject_code, False, f"找不到上年底稿: {subject_code}", None, []))
            continue

        # 处理单个科目
        success, message, output_path, warnings_list = process_single_subject(
            subject_code, template_path, prior_path, pmte_path,
            company_name, bs_date, output_dir, subject_config,
            functional_currency=functional_currency,
            accounting_standard=accounting_standard
        )

        results.append((subject_code, success, message, output_path, warnings_list))

    return results


if __name__ == "__main__":
    # 测试代码
    print("Roll Forward Core Module v4.0")
    print("请通过GUI或命令行调用此模块")
