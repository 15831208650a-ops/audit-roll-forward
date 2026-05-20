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
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

# 忽略openpyxl的警告
warnings.filterwarnings('ignore', category=UserWarning)

class RollForwardError(Exception):
    """Roll Forward 自定义异常"""
    pass


class SubjectConfig:
    """科目配置管理器"""

    def __init__(self, config_path="subjects_config.json"):
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
    pattern = config.get("prior_file_pattern", f"{subject_code}*")

    for search_dir in search_dirs:
        if not os.path.exists(search_dir):
            continue

        # 1. 尝试精确匹配：科目代码 + 年份
        for filename in os.listdir(search_dir):
            if filename.endswith('.xlsx') and not filename.startswith('~'):
                if subject_code in filename and str(prior_year) in filename:
                    return os.path.join(search_dir, filename)

        # 2. 放宽条件：只匹配科目代码
        for filename in os.listdir(search_dir):
            if filename.endswith('.xlsx') and not filename.startswith('~'):
                if subject_code in filename:
                    return os.path.join(search_dir, filename)

        # 3. 尝试用pattern中的关键字匹配（科目代码可能和文件名不完全一致）
        for filename in os.listdir(search_dir):
            if filename.endswith('.xlsx') and not filename.startswith('~'):
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
            if cell_company and company_name in str(cell_company):
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
    # 统一日期格式为 YYYY/MM/DD
    try:
        # 尝试解析日期
        date_obj = datetime.datetime.strptime(bs_date, "%Y-%m-%d")
        formatted_date = date_obj.strftime("%Y/%m/%d")
    except:
        formatted_date = bs_date

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
                if ws_new.cell(row=row, column=col + 1).value is None or "202" in str(ws_new.cell(row=row, column=col + 1).value):
                    ws_new.cell(row=row, column=col + 1, value=formatted_date)
                break
        else:
            continue
        break

    # 分析日期
    for row in [2, 3, 4, 5]:
        for col in [2, 3]:
            cell_val = ws_new.cell(row=row, column=col).value
            if cell_val and "分析日期" in str(cell_val):
                ws_new.cell(row=row, column=col + 1, value=datetime.datetime.now().strftime("%Y-%m-%d"))
                break
        else:
            continue
        break

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
    closing_col = lead_config.get("closing_col", 9)
    opening_col = lead_config.get("opening_col", 10)

    # 查找表头行
    header_search_text = lead_config.get("header_search_text", "期末审定数")
    header_row = find_header_row(ws_prior_values, header_search_text, (1, 80))

    if not header_row:
        return 0

    # 2.1 收集旧底稿的数据行（表头行下方到"合计"行之间的行）
    prior_data_rows = []  # [(row_num, c2_value, c3_value, c4_value, closing_value), ...]
    prior_header_row_new = header_row  # 旧底稿中的表头行

    for row in range(header_row + 1, ws_prior_values.max_row + 1):
        # 检测是否为"合计"行或数据结束标志
        c2_val = ws_prior_values.cell(row=row, column=2).value
        c5_val = ws_prior_values.cell(row=row, column=5).value

        # 如果遇到"合计"字样，认为是合计行，停止收集
        if c5_val and "合计" in str(c5_val):
            break

        # 收集有数据的行（至少C2或C4有值）
        c3_val = ws_prior_values.cell(row=row, column=3).value
        c4_val = ws_prior_values.cell(row=row, column=4).value
        closing_val = ws_prior_values.cell(row=row, column=closing_col).value

        # 如果有至少一个值，则认为是数据行
        if c2_val or c3_val or c4_val or closing_val is not None:
            prior_data_rows.append((row, c2_val, c3_val, c4_val, closing_val))

    # 2.2 确定模板中预留的数据行
    # 查找新模板中的"合计"行，确定预留数据行的范围
    new_data_start_row = None
    new_total_row = None

    for row in range(header_row + 1, ws_new.max_row + 1):
        c5_val = ws_new.cell(row=row, column=5).value
        if c5_val and "合计" in str(c5_val):
            new_total_row = row
            new_data_start_row = header_row + 1
            break

    # 2.2.1 精确计算模板中实际有数据的数据行数
    # 模板中表头行到合计行之间的行可能有空行，需要排除
    template_actual_data_rows = 0
    if new_data_start_row and new_total_row:
        for row in range(new_data_start_row, new_total_row):
            # 检查该行是否有实际数据（C2、C3、C4或C10有值）
            has_data = False
            for col in [2, 3, 4, 5, 10]:
                val = ws_new.cell(row=row, column=col).value
                if val is not None and str(val).strip() != '':
                    has_data = True
                    break
            if has_data:
                template_actual_data_rows += 1

    # 2.3 如果旧底稿数据行多于模板预留行数，需要插入行
    new_total_row_adjusted = new_total_row  # 记录插入行后的合计行位置
    if new_data_start_row and new_total_row:
        # 使用实际数据行数而非简单行数差
        template_data_row_count = template_actual_data_rows if 'template_actual_data_rows' in dir() else (new_total_row - new_data_start_row)
        prior_data_count = len(prior_data_rows)

        if prior_data_count > template_data_row_count:
            # 需要插入的行数
            rows_to_insert = prior_data_count - template_data_row_count
            # 在合计行前插入行
            insert_before_row = new_total_row
            for _ in range(rows_to_insert):
                ws_new.insert_rows(insert_before_row)
                # 复制上一行的格式到新插入的行（不复制公式，只复制值和格式）
                for col in range(1, ws_new.max_column + 1):
                    source_cell = ws_new.cell(row=insert_before_row - 1, column=col)
                    target_cell = ws_new.cell(row=insert_before_row, column=col)
                    # 复制非公式内容
                    if source_cell.value and not str(source_cell.value).startswith('='):
                        target_cell.value = source_cell.value

            # 更新合计行位置
            new_total_row_adjusted = new_total_row + rows_to_insert

            # 更新合计公式中的SUM范围
            # 插入行后，SUM范围需要扩展以包含新插入的行
            import re
            for col in range(1, ws_new.max_column + 1):
                cell = ws_new.cell(row=new_total_row_adjusted, column=col)
                if cell.value and str(cell.value).startswith('=') and 'SUM' in str(cell.value):
                    formula = str(cell.value)
                    # 扩展SUM范围：例如 SUM(F37:F40) -> SUM(F37:F41)
                    def expand_sum_range(match):
                        full = match.group(0)
                        if ':' in full:
                            # 提取范围部分
                            range_part = full.split('(')[1].split(')')[0]
                            start_ref, end_ref = range_part.split(':')
                            # 更新结束行号
                            end_col = re.sub(r'\d+', '', end_ref)
                            end_row_num = int(re.search(r'\d+', end_ref).group())
                            new_end_row = end_row_num + rows_to_insert
                            return f"SUM({start_ref}:{end_col}{new_end_row})"
                        return full
                    adjusted = re.sub(r'SUM\([^)]+\)', expand_sum_range, formula)
                    cell.value = adjusted

    # 2.4 复制数据：账套名称、科目编码、科目名称、期初数
    copied = 0
    current_row = header_row + 1
    for idx, (prior_row, c2_val, c3_val, c4_val, closing_val) in enumerate(prior_data_rows):
        target_row = header_row + 1 + idx

        # 复制C2（账套名称）
        if c2_val is not None:
            ws_new.cell(row=target_row, column=2).value = c2_val

        # 复制C3（科目编码）
        if c3_val is not None:
            ws_new.cell(row=target_row, column=3).value = c3_val

        # 复制C4（科目名称）
        if c4_val is not None:
            ws_new.cell(row=target_row, column=4).value = c4_val

        # 复制期初数（期末审定数 -> 期初审定数）
        if closing_val is not None:
            ws_new.cell(row=target_row, column=opening_col).value = closing_val
            copied += 1

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
        if prior_cell.value is not None:
            new_cell.value = prior_cell.value

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


def generate_output_filename(template_name, bs_date, company_name):
    """生成输出文件名"""
    output_name = template_name
    # 替换各种日期占位符
    output_name = re.sub(r'202[A-Za-z]{4,6}', bs_date.replace('-', ''), output_name)
    output_name = re.sub(r'\d{4}-\d{2}-\d{2}', bs_date, output_name)
    output_name = re.sub(r'\d{8}', bs_date.replace('-', ''), output_name)

    # 替换公司名称占位符
    output_name = re.sub(r'XYZ公司', company_name, output_name)

    return output_name


def process_single_subject(subject_code, template_path, prior_path, pmte_path,
                           company_name, bs_date, output_dir, subject_config,
                           cra_path=None):
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

            # 5. 处理Lead Sheet
            lead_sheet_name = lead_config.get("sheet_name", "")
            if lead_sheet_name and lead_sheet_name in wb_new.sheetnames and lead_sheet_name in wb_prior_formula.sheetnames:
                ws_new_lead = wb_new[lead_sheet_name]
                ws_prior_formula_lead = wb_prior_formula[lead_sheet_name]
                ws_prior_values_lead = wb_prior_values[lead_sheet_name]
                process_lead_sheet(ws_prior_formula_lead, ws_new_lead, ws_prior_values_lead,
                                   company_info, bs_date, company_name, lead_config, warnings_list)

            # 6. 处理K.01
            if k01_config.get("has_k01", False):
                k01_sheet_name = k01_config.get("sheet_name", "")
                if k01_sheet_name and k01_sheet_name in wb_new.sheetnames and k01_sheet_name in wb_prior_values.sheetnames:
                    ws_new_k01 = wb_new[k01_sheet_name]
                    ws_prior_k01 = wb_prior_values[k01_sheet_name]
                    process_k01(ws_prior_k01, ws_new_k01, k01_config)

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
            wb_new.save(output_path)

            # 生成警告消息
            warning_msg = ""
            if warnings_list:
                warning_msg = "; ".join(warnings_list)

            return True, f"处理成功{(' - ' + warning_msg if warning_msg else '')}", output_path, warnings_list

        finally:
            wb_new.close()
            wb_prior_formula.close()
            wb_prior_values.close()

    except Exception as e:
        return False, f"处理失败: {str(e)}", None, warnings_list


def process_multiple_subjects(subject_codes, template_dir, prior_dir, pmte_path,
                              company_name, bs_date, output_dir, config_path="subjects_config.json"):
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
            company_name, bs_date, output_dir, subject_config
        )

        results.append((subject_code, success, message, output_path, warnings_list))

    return results


if __name__ == "__main__":
    # 测试代码
    print("Roll Forward Core Module v4.0")
    print("请通过GUI或命令行调用此模块")
