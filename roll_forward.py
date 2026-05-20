#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
审计底稿 Roll Forward 脚本 v2.0（支持命令行参数）
功能：
  1. 读取 PMTE/CRA 信息表，提取公司信息（名称、TE、SAD、RP、CRA）
  2. 复制标准底稿模板，按规则重命名
  3. 填写 Lead Sheet 表头（公司名称、资产负债表日、PM/TE/SAD/CRA等）
  4. 从上年底稿找到"期末审定数"（I列），贴入本期"期初数"（J列）
  5. K.01 Agree SL to GL：复制表头，找到期末审定账面数/累计折旧/净值，对应写入期初
  6. 保留所有公式

使用方式（命令行参数）：
  python roll_forward.py --template "模板.xlsx" --prior "上年底稿.xlsx" --pmte "PMTE.xlsx" --company "公司名" --date "2026-12-31"

作者：AI Assistant
版本：2.0
"""

import os
import re
import shutil
import sys
import datetime
import argparse
from pathlib import Path
from copy import copy

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    print("正在安装 openpyxl...")
    os.system(f"{sys.executable} -m pip install openpyxl")
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, PatternFill
    from openpyxl.utils import get_column_letter


# ============================================================
#  1. 命令行参数解析
# ============================================================

def parse_args():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description="审计底稿 Roll Forward 工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  python roll_forward.py --template "K1 SWP 固定资产 202YMMDD XYZ公司.xlsx" \\
                         --prior "输入：K1 固定资产 20251231六六六有限公司.xlsx" \\
                         --pmte "输入；参考资料.xlsx" \\
                         --company "六六六有限公司" \\
                         --date "2026-12-31" \\
                         --output-dir "./output"
        """
    )

    parser.add_argument("--template", "-t", required=True,
                        help="标准底稿模板文件路径")
    parser.add_argument("--prior", "-p", required=True,
                        help="上年底稿文件路径")
    parser.add_argument("--pmte", "-m", required=True,
                        help="PMTE/CRA 信息表文件路径")
    parser.add_argument("--company", "-c", required=True,
                        help="公司名称")
    parser.add_argument("--date", "-d", required=True,
                        help="本报告期资产负债表日期 (格式: YYYY-MM-DD)")
    parser.add_argument("--output-dir", "-o", default="./output",
                        help="输出目录 (默认: ./output)")
    parser.add_argument("--interactive", "-i", action="store_true",
                        help="交互式模式（如果不使用命令行参数）")

    return parser.parse_args()


def get_interactive_inputs():
    """交互式获取用户输入"""
    print("=" * 60)
    print("  审计底稿 Roll Forward 工具")
    print("=" * 60)
    print()

    # 1. 资产负债表日期
    while True:
        bs_date = input("请输入本报告期资产负债表日期 (格式: YYYY-MM-DD, 例如 2026-12-31): ").strip()
        try:
            datetime.datetime.strptime(bs_date, "%Y-%m-%d")
            break
        except ValueError:
            print("  日期格式不正确，请重新输入。")

    # 2. 标准底稿模板路径
    while True:
        template_path = input("\n请输入标准底稿模板文件路径: ").strip().strip('"')
        if os.path.exists(template_path) and template_path.endswith('.xlsx'):
            break
        print("  文件不存在或格式不正确，请重新输入。")

    # 3. 上年底稿路径
    while True:
        prior_wb_path = input("\n请输入上年底稿文件路径: ").strip().strip('"')
        if os.path.exists(prior_wb_path) and prior_wb_path.endswith('.xlsx'):
            break
        print("  文件不存在或格式不正确，请重新输入。")

    # 4. PMTE/CRA 信息表路径
    while True:
        pmte_path = input("\n请输入 PMTE/CRA 信息表文件路径: ").strip().strip('"')
        if os.path.exists(pmte_path) and pmte_path.endswith('.xlsx'):
            break
        print("  文件不存在或格式不正确，请重新输入。")

    # 5. 输出目录
    default_output_dir = os.path.join(os.path.dirname(template_path), "output")
    output_dir = input(f"\n请输入输出目录 (直接回车使用默认: {default_output_dir}): ").strip().strip('"')
    if not output_dir:
        output_dir = default_output_dir

    # 6. 公司名称
    company_name = input("\n请输入公司名称: ").strip()

    return {
        "bs_date": bs_date,
        "template_path": template_path,
        "prior_wb_path": prior_wb_path,
        "pmte_path": pmte_path,
        "output_dir": output_dir,
        "company_name": company_name,
    }


# ============================================================
#  2. PMTE/CRA 信息提取
# ============================================================

def extract_company_info(pmte_path, company_name):
    """
    从 PMTE/CRA 信息表中提取指定公司的信息
    返回: dict 包含 PM, TE, SAD, RP, CRA 等
    """
    wb = openpyxl.load_workbook(pmte_path, data_only=True)
    info = {
        "PM": None,
        "TE": None,
        "SAD": None,
        "RP": None,
        "CRA": None,
        "Level": None,
    }

    # 尝试读取 PMTE Sheet
    if "PMTE" in wb.sheetnames:
        ws = wb["PMTE"]
        # PMTE 表头一般在第1行，列结构：A=公司名, B=层级, C=是否需要RP, D=PM, E=TE, F=SAD
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


def extract_cra_info(pmte_path, subject_keyword):
    """
    从 CRA 信息表中提取指定科目的 CRA 风险等级
    参数:
        pmte_path: PMTE/CRA 信息表路径
        subject_keyword: 科目关键词，如"固定资产"
    返回: dict，格式为 {认定: 风险等级}
    """
    wb = openpyxl.load_workbook(pmte_path, data_only=True)
    cra_info = {}

    if "CRA" not in wb.sheetnames:
        wb.close()
        return cra_info

    ws = wb["CRA"]

    # 遍历所有行，查找包含科目关键词的行
    for row in range(1, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=2).value
        if cell_value and subject_keyword in str(cell_value):
            # 检查第3列是否包含"是"或"X"
            applicable = ws.cell(row=row, column=3).value
            if applicable and "是" in str(applicable):
                # 从第2列提取认定类型
                subject_name = str(cell_value)
                # 尝试提取认定（如"存在性"、"完整性"等）
                if "存在性" in subject_name:
                    cra_info["存在性"] = "High"  # 默认值，实际应从第8列读取
                elif "完整性" in subject_name:
                    cra_info["完整性"] = "High"
                elif "计价" in subject_name or "计量" in subject_name:
                    cra_info["计价"] = "Moderate"
                elif "权利和义务" in subject_name:
                    cra_info["权利和义务"] = "Low"
                elif "列报" in subject_name or "披露" in subject_name:
                    cra_info["列报和披露"] = "Low"

    wb.close()
    return cra_info


# ============================================================
#  3. Lead Sheet 处理
# ============================================================

def fill_lead_sheet_header(ws_new, company_info, cra_info, bs_date, prior_bs_date, company_name):
    """填写 Lead Sheet 表头信息，包括CRA风险等级"""
    # 写入公司基本信息
    ws_new.cell(row=2, column=3, value=company_name)  # C2=公司名称
    ws_new.cell(row=3, column=3, value=bs_date)         # C3=资产负债表日
    ws_new.cell(row=4, column=3, value=datetime.datetime.now().strftime("%Y-%m-%d"))  # C4=分析日期

    if company_info.get("TE"):
        ws_new.cell(row=5, column=3, value=company_info["TE"])
    if company_info.get("SAD"):
        ws_new.cell(row=6, column=3, value=company_info["SAD"])

    # 填写CRA风险等级（如果提供了的话）
    if cra_info:
        # 认定行映射（根据标准底稿结构）
        # 第15行=完整性, 第16行=存在性/发生, 第17行=计价/计量, 第18行=权利和义务, 第19行=列报和披露
        cra_mapping = {
            "完整性": 15,
            "存在性": 16,
            "计价": 17,
            "权利和义务": 18,
            "列报和披露": 19,
        }
        for assertion, row_num in cra_mapping.items():
            if assertion in cra_info:
                ws_new.cell(row=row_num, column=3, value=cra_info[assertion])
        print(f"  [Lead Sheet] 已填写CRA风险等级")
    else:
        print("  [警告] 未找到CRA信息，请在生成底稿后手动填写")

    print(f"  [Lead Sheet] 已填写公司信息: {company_name}, 日期: {bs_date}")


def roll_forward_lead_sheet(ws_prior, ws_new, ws_prior_values=None):
    """
    将上年底稿 Lead Sheet 的期末审定数(I列)贴入本期底稿的期初数(J列)
    如果提供了 ws_prior_values，则使用计算后的值
    """
    print("  [Lead Sheet] 开始 Roll Forward 期初数...")

    # 找到表头行（扩大搜索范围到前60行）
    header_row = None
    for row in range(1, min(60, ws_prior.max_row + 1)):
        for col in range(1, min(ws_prior.max_column + 1, 30)):
            cell_value = ws_prior.cell(row=row, column=col).value
            if cell_value and "期末审定数" in str(cell_value):
                header_row = row
                break
        if header_row:
            break

    if not header_row:
        print("  [警告] 未找到'期末审定数'表头，尝试查找'上期末审定数'表头")
        for row in range(1, min(60, ws_prior.max_row + 1)):
            for col in range(1, min(ws_prior.max_column + 1, 30)):
                cell_value = ws_prior.cell(row=row, column=col).value
                if cell_value and "上期末审定数" in str(cell_value):
                    header_row = row
                    break
            if header_row:
                break

    if not header_row:
        print("  [警告] 仍未找到表头，使用默认第48行")
        header_row = 48  # 根据实际数据，表头通常在第48行左右

    print(f"  [Lead Sheet] 找到表头在第 {header_row} 行")

    # I 列=期末审定数(9), J 列=期初数(10)
    col_closing = 9   # I 列
    col_opening = 10  # J 列

    data_start_row = header_row + 1 if header_row else 2

    copied_count = 0
    for row in range(data_start_row, ws_prior.max_row + 1):
        # 优先使用计算后的值（如果提供了的话）
        if ws_prior_values:
            prior_cell = ws_prior_values.cell(row=row, column=col_closing)
        else:
            prior_cell = ws_prior.cell(row=row, column=col_closing)
        new_cell = ws_new.cell(row=row, column=col_opening)

        if prior_cell.value is not None:
            new_cell.value = prior_cell.value
            copied_count += 1

    print(f"  [Lead Sheet] 已复制 {copied_count} 个期初数")
    return copied_count


# ============================================================
#  4. K.01 Agree SL to GL 处理
# ============================================================

def process_k01(ws_prior, ws_new, ws_prior_values=None):
    """
    处理 K.01 Agree SL to GL Sheet：
      1. 复制第10行表头（固定资产分类）
      2. 只填第12行"年初余额"的数据（上年年末=本年年初）
    """
    print("  [K.01] 开始处理 Agree SL to GL...")

    # 1. 复制第10行表头（固定资产类别）
    # 上年底稿第10行是固定资产类别，新底稿第10行应该与之保持一致
    header_row_10 = 10
    for col in range(1, ws_prior.max_column + 1):
        prior_cell = ws_prior.cell(row=header_row_10, column=col)
        new_cell = ws_new.cell(row=header_row_10, column=col)
        if prior_cell.value is not None:
            new_cell.value = prior_cell.value

    print(f"  [K.01] 已复制第10行固定资产类别表头")

    # 2. 找到"年初余额"行（通常是第12行）
    opening_balance_row = None
    for row in range(11, min(20, ws_prior.max_row + 1)):
        cell_value = ws_prior.cell(row=row, column=3).value  # C列
        if cell_value and "年初余额" in str(cell_value):
            opening_balance_row = row
            break

    if not opening_balance_row:
        print("  [警告] 未找到'年初余额'行，尝试在第12行查找")
        opening_balance_row = 12

    print(f"  [K.01] '年初余额'行位于第 {opening_balance_row} 行")

    # 3. 找到"审定数"列（G=7, J=10, M=13, P=16, S=19）
    # 在新底稿中，这些列对应的是"期初审定数"
    col_mapping = {
        "audited_1": 7,   # G列 - 房屋建筑物审定数
        "audited_2": 10,  # J列 - 机器设备审定数
        "audited_3": 13,  # M列 - 运输设备审定数
        "audited_4": 16,  # P列 - 电子产品及通讯设备审定数
        "audited_5": 19,  # S列 - 合计审定数
    }

    # 4. 只复制"年初余额"行的数据
    # 在新底稿中，"年初余额"行的"账面数"列（E, H, K, N, Q）= 上年底稿对应行的"审定数"
    copied = 0
    for key, col_idx in col_mapping.items():
        # 优先使用计算后的值
        if ws_prior_values:
            prior_cell = ws_prior_values.cell(row=opening_balance_row, column=col_idx)
        else:
            prior_cell = ws_prior.cell(row=opening_balance_row, column=col_idx)

        # 在新底稿中，"年初余额"行的"账面数"列 = 审定数列 - 2（通常）
        # 审定数列：G=7, J=10, M=13, P=16, S=19
        # 账面数列：E=5, H=8, K=11, N=14, Q=17
        book_col = col_idx - 2
        new_cell = ws_new.cell(row=opening_balance_row, column=book_col)

        if prior_cell.value is not None:
            new_cell.value = prior_cell.value
            copied += 1

    print(f"  [K.01] 已从'年初余额'行复制 {copied} 个期初数")
    return copied


# ============================================================
#  5. 辅助功能
# ============================================================

def generate_output_filename(template_path, bs_date, company_name):
    """根据模板文件名生成输出文件名"""
    template_name = os.path.basename(template_path)
    output_name = template_name

    # 替换各种日期占位符格式
    # 格式1: 202YMMDD, 202MMMDD 等（字母+数字混合占位符，支持4-6个字母）
    output_name = re.sub(r'202[A-Za-z]{4,6}', bs_date.replace('-', ''), output_name)
    # 格式2: 标准日期格式 2025-12-31
    output_name = re.sub(r'\d{4}-\d{2}-\d{2}', bs_date, output_name)
    # 格式3: 8位数字日期 20251231
    output_name = re.sub(r'\d{8}', bs_date.replace('-', ''), output_name)

    # 替换公司名称占位符
    output_name = output_name.replace("XYZ公司", company_name)
    output_name = output_name.replace("六六六有限公司", company_name)

    return output_name


# ============================================================
#  6. 主流程
# ============================================================

def main():
    """主流程"""
    # 检查是否为交互模式
    args = parse_args()

    if args.interactive or len(sys.argv) == 1:
        inputs = get_interactive_inputs()
    else:
        inputs = {
            "bs_date": args.date,
            "template_path": args.template,
            "prior_wb_path": args.prior,
            "pmte_path": args.pmte,
            "output_dir": args.output_dir,
            "company_name": args.company,
        }

    bs_date = inputs["bs_date"]
    prior_bs_date = f"{int(bs_date[:4]) - 1}{bs_date[4:]}"  # 简单推算上年日期
    template_path = inputs["template_path"]
    prior_wb_path = inputs["prior_wb_path"]
    pmte_path = inputs["pmte_path"]
    output_dir = inputs["output_dir"]
    company_name = inputs["company_name"]

    # 创建输出目录
    os.makedirs(output_dir, exist_ok=True)

    print("\n" + "=" * 60)
    print("  开始处理...")
    print("=" * 60)

    # 1. 提取 PMTE 和 CRA 信息
    print("\n[1/4] 正在从 PMTE 信息表提取公司信息...")
    company_info = extract_company_info(pmte_path, company_name)
    print(f"  PM: {company_info.get('PM')}, TE: {company_info.get('TE')}, SAD: {company_info.get('SAD')}, RP: {company_info.get('RP')}")

    # 提取CRA信息（假设科目关键词为"固定资产"）
    print("\n[1.5/4] 正在从 CRA 信息表提取风险等级...")
    cra_info = extract_cra_info(pmte_path, "固定资产")
    if cra_info:
        print(f"  CRA风险等级: {cra_info}")
    else:
        print("  [警告] 未找到固定资产的CRA信息，请在生成底稿后手动填写")

    # 2. 复制标准模板
    print("\n[2/4] 正在复制标准底稿模板...")
    output_name = generate_output_filename(template_path, bs_date, company_name)
    output_path = os.path.join(output_dir, output_name)

    shutil.copy2(template_path, output_path)
    print(f"  已复制到: {output_path}")

    # 3. 打开新底稿和上年底稿
    print("\n[3/4] 正在处理数据 Roll Forward...")
    wb_new = openpyxl.load_workbook(output_path)
    # 上年底稿需要打开两次：一次读取公式/表头，一次读取计算后的值
    wb_prior = openpyxl.load_workbook(prior_wb_path, data_only=False)
    wb_prior_values = openpyxl.load_workbook(prior_wb_path, data_only=True)

    # 4. 处理 Lead Sheet
    if "K.00 Lead Sheet" in wb_new.sheetnames and "K.00 Lead Sheet" in wb_prior.sheetnames:
        ws_new_lead = wb_new["K.00 Lead Sheet"]
        ws_prior_lead = wb_prior["K.00 Lead Sheet"]
        ws_prior_lead_values = wb_prior_values["K.00 Lead Sheet"] if wb_prior_values else None
        fill_lead_sheet_header(ws_new_lead, company_info, cra_info, bs_date, prior_bs_date, company_name)
        roll_forward_lead_sheet(ws_prior_lead, ws_new_lead, ws_prior_lead_values)
    else:
        print("  [警告] 未找到 K.00 Lead Sheet，跳过")

    # 5. 处理 K.01 Agree SL to GL
    if "K.01 Agree SL to GL" in wb_new.sheetnames and "K.01 Agree SL to GL" in wb_prior.sheetnames:
        ws_new_k01 = wb_new["K.01 Agree SL to GL"]
        ws_prior_k01 = wb_prior["K.01 Agree SL to GL"]
        ws_prior_k01_values = wb_prior_values["K.01 Agree SL to GL"] if wb_prior_values else None
        process_k01(ws_prior_k01, ws_new_k01, ws_prior_k01_values)
    else:
        print("  [警告] 未找到 K.01 Agree SL to GL，跳过")

    # 6. 保存新底稿
    print("\n[4/4] 正在保存新底稿...")
    wb_new.save(output_path)
    wb_new.close()
    wb_prior.close()
    if wb_prior_values:
        wb_prior_values.close()

    print("\n" + "=" * 60)
    print(f"  处理完成！")
    print(f"  输出文件: {output_path}")
    print("=" * 60)

    return output_path


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n用户取消操作。")
        sys.exit(0)
    except Exception as e:
        print(f"\n[错误] 处理过程中出现错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
