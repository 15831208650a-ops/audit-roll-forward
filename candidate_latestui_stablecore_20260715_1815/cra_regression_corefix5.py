import json
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication

from cra_support import CRA_PARSER_VERSION, match_subject, parse_cra_paste_text
from main_gui import RollForwardApp


PASSES = []
FAILURES = []


def check(condition, message):
    (PASSES if condition else FAILURES).append(message)


EIGHT_COLUMN_TEXT = """帐户认定\t\t\t固有风险\t控制风险\t综合风险评估\t\t测试界限
C. 货币资金-存在性\t\t\t较低\t不依赖控制\t中等程度\tY\t50%
C. 货币资金-计价/计量\t\t\t较低\t不依赖控制\t中等程度\tY\t50%
VC. 销售费用-完整性\t\t\t较低\t不依赖控制\t中等程度\tY\t15%
VC. 销售费用-存在性/发生\t\t\t较低\t不依赖控制\t中等程度\tN\t15%
VD. 管理费用-完整性\t\t\t较低\t不依赖控制\t中等程度\tY\t15%
VD. 管理费用-计价/计量\t\t\t较低\t不依赖控制\t中等程度\tN/A\t15%
VD. 管理费用-列报和披露\t\t\t较低\t不依赖控制\t中等程度\t\tN/A"""


NINE_COLUMN_TEXT = """帐户认定\t\t\t\t固有风险\t控制风险\t综合风险评估\t\t测试界限
C\t货币资金\t\t\t\t\t\t\t
\tC. 货币资金-存在性\t\t\t较低\t不依赖控制\t中等程度\tY\t50%
\tC. 货币资金-计价/计量\t\t\t较低\t不依赖控制\t中等程度\tY\t50%
VC\t销售费用\t\t\t\t\t\t\t
\tVC. 销售费用-完整性\t\t\t较低\t不依赖控制\t中等程度\tY\t15%
\tVC. 销售费用-存在性/发生\t\t\t较低\t不依赖控制\t中等程度\tN\t15%
VD\t管理费用\t\t\t\t\t\t\t
\tVD. 管理费用-完整性\t\t\t较低\t不依赖控制\t中等程度\tY\t15%
\tVD. 管理费用-计价/计量\t\t\t较低\t不依赖控制\t中等程度\tN/A\t15%
\tVD. 管理费用-列报和披露\t\t\t较低\t不依赖控制\t中等程度\t\tN/A"""


def records_by_subject(text):
    records = parse_cra_paste_text(text, ["C", "UexpVCVD"])
    return records, {
        subject: [record for record in records if record.get("subject_code") == subject]
        for subject in ("C", "UexpVCVD")
    }


def check_parser_layout(text, label):
    records, grouped = records_by_subject(text)
    check(len(grouped["C"]) == 2, f"{label}: two cash assertions map to C")
    check(
        all(record["account_name"] == "货币资金" for record in grouped["C"]),
        f"{label}: C account name remains 货币资金",
    )
    check(len(grouped["UexpVCVD"]) == 5, f"{label}: VC and VD map to UexpVCVD")
    check(
        {record["account_name"] for record in grouped["UexpVCVD"]} == {"销售费用", "管理费用"},
        f"{label}: sales and administrative account names remain distinct",
    )
    not_applicable = [record for record in grouped["UexpVCVD"] if record.get("applicable") is False]
    check(len(not_applicable) == 2, f"{label}: N and N/A are not applicable")
    check(
        all(record["cra_level"] == "N/A" and record["ratio_text"] == "N/A" for record in not_applicable),
        f"{label}: not-applicable CRA and ratio display N/A",
    )
    blank_applicable = next(record for record in grouped["UexpVCVD"] if record["assertion"] == "P&D")
    check(
        blank_applicable["applicable"] is True and blank_applicable["cra_level"] == "Moderate",
        f"{label}: blank applicability remains applicable even when ratio is N/A",
    )
    return records


def check_preview(records):
    app = QApplication.instance() or QApplication([])
    window = RollForwardApp()
    window.save_workbench_data = lambda: None

    legacy_company = window.create_empty_company("Fixture")
    legacy_company.update({
        "cra_text": NINE_COLUMN_TEXT,
        "cra_table_records": records,
        "cra_parser_version": "legacy-parser",
        "apply_cra": True,
        "subjects": ["C", "UexpVCVD"],
    })
    window.project_data = {"project_name": "Fixture", "companies": [legacy_company]}
    window.current_company_index = 0
    window.load_company_to_form(0)
    check(window.cra_records_stale is True, "Preview: records from an older parser are marked stale")
    check(window.cra_table.rowCount() == len(records), "Preview: legacy records remain visible for comparison")
    check(window.apply_cra_checkbox.isEnabled() is False, "Preview: legacy records cannot be written before reparse")

    window.cra_records_stale = False
    legacy_company["cra_parser_version"] = CRA_PARSER_VERSION
    window.apply_cra_checkbox.setEnabled(True)
    window.populate_cra_table(records)

    subject_index = window.cra_subject_filter.findData("UexpVCVD")
    window.cra_subject_filter.setCurrentIndex(subject_index)
    visible_subjects = [
        window.cra_table_text(row, 1)
        for row in range(window.cra_table.rowCount())
        if not window.cra_table.isRowHidden(row)
    ]
    check(visible_subjects and set(visible_subjects) == {"UexpVCVD"}, "Preview: subject filter hides rows without deleting them")
    check(window.cra_table.rowCount() == len(records), "Preview: filtering retains all parsed rows")

    window.cra_subject_filter.setCurrentIndex(0)
    window.cra_filter_input.setText("货币资金")
    visible_accounts = [
        window.cra_table_text(row, 2)
        for row in range(window.cra_table.rowCount())
        if not window.cra_table.isRowHidden(row)
    ]
    check(visible_accounts and set(visible_accounts) == {"货币资金"}, "Preview: search filters account names")

    window.cra_filter_input.clear()
    window.cra_table.sortItems(5, Qt.SortOrder.AscendingOrder)
    numeric_values = [
        window.cra_table.item(row, 5).data(Qt.ItemDataRole.UserRole)
        for row in range(window.cra_table.rowCount())
    ]
    check(numeric_values == sorted(numeric_values), "Preview: ratio sorting is numeric")

    window.cra_records_stale = False
    window.apply_cra_checkbox.setEnabled(True)
    window.apply_cra_checkbox.setChecked(True)
    original_rows = window.cra_table.rowCount()
    window.cra_text_input.setPlainText(NINE_COLUMN_TEXT + "\n")
    check(window.cra_records_stale is True, "Preview: changed source text marks records stale")
    check(window.apply_cra_checkbox.isEnabled() is False, "Preview: stale records cannot be enabled for writing")
    check(window.collect_cra_records() == [], "Preview: stale records are never collected for workpaper writing")
    check(window.cra_table.rowCount() == original_rows, "Preview: stale table remains visible for comparison")
    window.close()
    app.processEvents()


def main():
    check(match_subject("VC") == "UexpVCVD", "Subject matching: bare VC maps to UexpVCVD")
    check(match_subject("VD") == "UexpVCVD", "Subject matching: bare VD maps to UexpVCVD")
    check(match_subject("VC&VD") == "UexpVCVD", "Subject matching: VC&VD maps to UexpVCVD")
    check(match_subject("C. 货币资金") == "C", "Subject matching: explicit C cash maps to C")
    check(match_subject("SC") != "C", "Subject matching: containing letter C does not imply cash")

    eight_records = check_parser_layout(EIGHT_COLUMN_TEXT, "8-column Canvas")
    nine_records = check_parser_layout(NINE_COLUMN_TEXT, "9-column Canvas")
    check(
        [(record["subject_code"], record["account_name"], record["assertion"], record["applicable"]) for record in eight_records]
        == [(record["subject_code"], record["account_name"], record["assertion"], record["applicable"]) for record in nine_records],
        "Canvas layouts: 8-column and 9-column forms produce equivalent records",
    )
    check_preview(nine_records)

    print(json.dumps({"passes": len(PASSES), "failures": FAILURES}, ensure_ascii=False, indent=2))
    if FAILURES:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
