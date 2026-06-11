#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Audit Roll Forward desktop application.

Technology stack: PyQt6 + openpyxl + PyInstaller
"""

import os
import sys
import json
import datetime
import subprocess
import tempfile
from pathlib import Path

from PyQt6.QtCore import QSize, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QDesktopServices, QFont
from PyQt6.QtCore import QUrl
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from roll_forward_core import SubjectConfig, process_multiple_subjects, resource_path


EY_YELLOW = "#FFE600"
EY_BLACK = "#000000"
EY_OFF_BLACK = "#2E2E38"
EY_PANEL = "#1A1A24"
EY_PANEL_ALT = "#23232F"
EY_BORDER = "#3A3A4A"
EY_TEXT = "#F6F6FA"
EY_MUTED = "#D7D7E2"
EY_PLACEHOLDER = "#9FA0B3"
EY_SUCCESS = "#2DBE60"
EY_ERROR = "#FF4B55"
FEEDBACK_URL = "https://v.wjx.cn/vm/mEfMFm4.aspx"
APP_STATE_DIR = Path(os.getenv("APPDATA", str(Path.home()))) / "AuditRollForward"
FEEDBACK_STATE_PATH = APP_STATE_DIR / "feedback_state.json"
APP_LOG_PATH = APP_STATE_DIR / "logs" / "app.log"


class RollForwardApp(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.setWindowTitle("Audit Roll Forward")
        self.setGeometry(80, 60, 1400, 1000)
        self.setMinimumSize(1100, 760)
        self.setStyleSheet(self.stylesheet())
        self.init_ui()

    def stylesheet(self):
        return f"""
            QWidget {{
                font-family: "Microsoft YaHei", "Segoe UI", Arial, sans-serif;
                font-size: 13px;
                color: {EY_TEXT};
            }}
            QMainWindow {{
                background: {EY_BLACK};
            }}
            QWidget#AppRoot {{
                background: {EY_BLACK};
            }}
            QWidget#Page {{
                background: {EY_BLACK};
            }}
            QWidget#Header {{
                background: {EY_BLACK};
                border: 1px solid {EY_BORDER};
                border-radius: 4px;
            }}
            QScrollArea {{
                border: none;
                background: {EY_BLACK};
            }}
            QScrollArea#RootScroll {{
                background: {EY_BLACK};
            }}
            QScrollArea#RootScroll > QWidget {{
                background: {EY_BLACK};
            }}
            QFrame#Card {{
                background: {EY_PANEL};
                border: 1px solid {EY_BORDER};
                border-radius: 4px;
            }}
            QLabel#EyMark {{
                background: {EY_YELLOW};
                color: {EY_BLACK};
                font-size: 26px;
                font-weight: 800;
                padding: 14px 18px;
            }}
            QLabel#Title {{
                color: {EY_TEXT};
                font-size: 26px;
                font-weight: 700;
            }}
            QLabel#Subtitle {{
                color: {EY_MUTED};
                font-size: 12px;
            }}
            QLabel#SectionTitle {{
                color: {EY_TEXT};
                font-size: 15px;
                font-weight: 700;
                padding-bottom: 8px;
                border-bottom: 1px solid {EY_BORDER};
            }}
            QLabel#FieldLabel {{
                color: {EY_MUTED};
                font-size: 12px;
                font-weight: 600;
            }}
            QLineEdit {{
                background: {EY_OFF_BLACK};
                border: 1px solid {EY_BORDER};
                border-radius: 3px;
                color: {EY_TEXT};
                padding: 9px 11px;
                selection-background-color: {EY_YELLOW};
                selection-color: {EY_BLACK};
            }}
            QLineEdit:focus {{
                border: 1px solid {EY_YELLOW};
            }}
            QLineEdit[readOnly="true"] {{
                color: {EY_MUTED};
            }}
            QLineEdit::placeholder {{
                color: {EY_PLACEHOLDER};
            }}
            QPushButton {{
                background: {EY_PANEL_ALT};
                color: {EY_TEXT};
                border: 1px solid {EY_BORDER};
                border-radius: 3px;
                padding: 9px 14px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                border-color: {EY_YELLOW};
                color: {EY_YELLOW};
            }}
            QPushButton:disabled {{
                color: #737383;
                border-color: #333340;
                background: #15151D;
            }}
            QPushButton#PrimaryButton {{
                background: {EY_YELLOW};
                color: {EY_BLACK};
                border: 1px solid {EY_YELLOW};
                padding: 13px 24px;
                font-size: 14px;
                font-weight: 800;
            }}
            QPushButton#PrimaryButton:hover {{
                background: #FFF166;
                color: {EY_BLACK};
            }}
            QPushButton#SecondaryButton {{
                color: {EY_YELLOW};
                border-color: {EY_YELLOW};
            }}
            QListWidget {{
                background: {EY_OFF_BLACK};
                border: 1px solid {EY_BORDER};
                border-radius: 3px;
                padding: 10px;
                outline: none;
            }}
            QListWidget::item {{
                padding: 10px 12px;
                margin: 4px;
                border: 1px solid {EY_BORDER};
                border-left: 4px solid transparent;
                border-radius: 3px;
                color: {EY_MUTED};
            }}
            QListWidget::item:hover {{
                background: #333344;
                border-color: {EY_YELLOW};
            }}
            QListWidget::item:selected {{
                background: #3A3520;
                border: 1px solid {EY_YELLOW};
                border-left: 4px solid {EY_YELLOW};
                color: {EY_TEXT};
            }}
            QCheckBox {{
                spacing: 8px;
                color: {EY_TEXT};
            }}
            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
                border: 1px solid {EY_BORDER};
                background: {EY_OFF_BLACK};
            }}
            QCheckBox::indicator:checked {{
                background: {EY_YELLOW};
                border-color: {EY_YELLOW};
            }}
            QCheckBox#SubjectCheckbox {{
                background: {EY_OFF_BLACK};
                border: 1px solid {EY_BORDER};
                border-radius: 3px;
                padding: 10px 12px;
                min-height: 32px;
                font-weight: 600;
            }}
            QCheckBox#SubjectCheckbox:hover {{
                border-color: {EY_YELLOW};
                color: {EY_YELLOW};
            }}
            QCheckBox#SubjectCheckbox:checked {{
                background: #3A3520;
                border-color: {EY_YELLOW};
                color: {EY_TEXT};
            }}
            QGroupBox {{
                border: 1px solid {EY_BORDER};
                border-radius: 3px;
                margin-top: 12px;
                padding: 16px 12px 12px 12px;
                color: {EY_MUTED};
                font-weight: 600;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }}
            QProgressBar {{
                background: {EY_OFF_BLACK};
                border: 1px solid {EY_BORDER};
                border-radius: 3px;
                color: {EY_TEXT};
                height: 24px;
                text-align: center;
                font-weight: 700;
            }}
            QProgressBar::chunk {{
                background: {EY_YELLOW};
            }}
            QTextEdit {{
                background: {EY_OFF_BLACK};
                border: 1px solid {EY_BORDER};
                border-radius: 3px;
                color: {EY_MUTED};
                padding: 10px;
                font-family: Consolas, "Courier New", monospace;
                font-size: 12px;
            }}
            QTableWidget {{
                background: {EY_OFF_BLACK};
                border: 1px solid {EY_BORDER};
                gridline-color: {EY_BORDER};
                selection-background-color: #3A3520;
                selection-color: {EY_TEXT};
            }}
            QHeaderView::section {{
                background: {EY_PANEL_ALT};
                color: {EY_YELLOW};
                border: 1px solid {EY_BORDER};
                padding: 7px;
                font-weight: 700;
            }}
            QScrollBar:vertical {{
                background: {EY_BLACK};
                width: 8px;
            }}
            QScrollBar::handle:vertical {{
                background: {EY_BORDER};
                min-height: 30px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: {EY_YELLOW};
            }}
        """

    def init_ui(self):
        self.setObjectName("AppRoot")
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setObjectName("RootScroll")
        scroll.setWidgetResizable(True)
        scroll.viewport().setStyleSheet(f"background: {EY_BLACK};")
        root_layout.addWidget(scroll)

        container = QWidget()
        container.setObjectName("Page")
        container.setStyleSheet(f"background: {EY_BLACK};")
        scroll.setWidget(container)
        layout = QVBoxLayout(container)
        layout.setContentsMargins(34, 28, 34, 28)
        layout.setSpacing(18)

        layout.addWidget(self.create_header())
        layout.addWidget(self.create_parameters_card())
        layout.addWidget(self.create_file_card())
        layout.addWidget(self.create_subject_card())
        layout.addWidget(self.create_action_card())
        layout.addWidget(self.create_log_card())
        layout.addWidget(self.create_results_card())
        layout.addWidget(self.create_footer())

    def create_header(self):
        header = QWidget()
        header.setObjectName("Header")
        header.setStyleSheet(
            f"QWidget#Header {{ background: {EY_BLACK}; border: 1px solid {EY_BORDER}; border-radius: 4px; }}"
        )
        layout = QHBoxLayout(header)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(16)

        mark = QLabel("EY")
        mark.setObjectName("EyMark")
        mark.setFixedSize(76, 76)
        mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        mark.setStyleSheet(
            f"background: {EY_YELLOW}; color: {EY_BLACK}; "
            "font-size: 26px; font-weight: 800;"
        )
        layout.addWidget(mark)

        title_box = QVBoxLayout()
        title_box.setSpacing(3)
        title = QLabel("Audit Roll Forward")
        title.setObjectName("Title")
        subtitle = QLabel("底稿自动结转工具 | 标准模板、上年底稿、PMTE 参数表一键生成")
        subtitle.setObjectName("Subtitle")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)
        layout.addLayout(title_box)
        layout.addStretch()

        badge = QLabel("TEST BUILD")
        badge.setStyleSheet(
            f"color: {EY_BLACK}; background: {EY_YELLOW}; padding: 6px 10px; "
            "font-weight: 800; border-radius: 3px;"
        )
        badge.setMinimumHeight(42)
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(badge)
        return header

    def create_card(self, title):
        card = QFrame()
        card.setObjectName("Card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 16, 18, 18)
        layout.setSpacing(14)

        title_label = QLabel(title)
        title_label.setObjectName("SectionTitle")
        layout.addWidget(title_label)
        return card

    def create_parameters_card(self):
        card = self.create_card("基础参数")
        grid = QGridLayout()
        grid.setHorizontalSpacing(16)
        grid.setVerticalSpacing(12)
        grid.setColumnStretch(1, 3)
        grid.setColumnStretch(3, 3)

        self.company_input = self.create_input("例如：东风鸿远工程咨询有限公司")
        self.date_input = self.create_input("例如：2026/12/31 或 20261231")
        self.functional_currency_input = self.create_input("例如：人民币")
        self.accounting_standard_input = self.create_input("例如：企业会计准则")
        self.pm_input = self.create_input("可选：手工输入 PM")
        self.te_input = self.create_input("可选：手工输入 TE")

        grid.addWidget(self.create_label("公司名称"), 0, 0)
        grid.addWidget(self.company_input, 0, 1)
        grid.addWidget(self.create_label("资产负债表日"), 0, 2)
        grid.addWidget(self.date_input, 0, 3)
        grid.addWidget(self.create_label("记账本位币"), 1, 0)
        grid.addWidget(self.functional_currency_input, 1, 1)
        grid.addWidget(self.create_label("适用会计准则"), 1, 2)
        grid.addWidget(self.accounting_standard_input, 1, 3)
        grid.addWidget(self.create_label("PM"), 2, 0)
        grid.addWidget(self.pm_input, 2, 1)
        grid.addWidget(self.create_label("TE"), 2, 2)
        grid.addWidget(self.te_input, 2, 3)

        option_box = QGroupBox("处理选项")
        option_layout = QGridLayout(option_box)
        option_layout.setHorizontalSpacing(18)
        option_layout.setVerticalSpacing(10)

        self.roll_wording_checkbox = QCheckBox("Roll forward wording / 分析说明 / 调整分录汇总")
        self.roll_wording_checkbox.setToolTip("勾选后会复制上年底稿中的说明文字和调整汇总，并用黄色标注需复核区域。")
        self.generate_summary_checkbox = QCheckBox("生成 Roll Forward Summary")
        self.generate_summary_checkbox.setChecked(True)
        self.generate_summary_checkbox.setToolTip("每个输出底稿附加检查报告工作表，列示更新、标黄和未匹配信息。")
        self.keep_adjustments_checkbox = QCheckBox("保留上年调整分录")
        self.keep_adjustments_checkbox.setToolTip("测试选项：用于记录测试人员希望保留上年调整分录的场景。")
        self.clear_testing_checkbox = QCheckBox("清空本期测试区域")
        self.clear_testing_checkbox.setToolTip("测试选项：用于记录测试人员希望清空本期测试区域的场景。")

        option_layout.addWidget(self.roll_wording_checkbox, 0, 0)
        option_layout.addWidget(self.generate_summary_checkbox, 0, 1)
        option_layout.addWidget(self.keep_adjustments_checkbox, 1, 0)
        option_layout.addWidget(self.clear_testing_checkbox, 1, 1)
        grid.addWidget(option_box, 3, 0, 1, 4)

        card.layout().addLayout(grid)
        return card

    def create_file_card(self):
        card = self.create_card("文件路径")
        layout = QGridLayout()
        layout.setHorizontalSpacing(12)
        layout.setVerticalSpacing(12)
        layout.setColumnStretch(1, 1)

        self.prior_dir_input = self.create_file_input("请选择上年底稿目录")
        self.pmte_input = self.create_file_input("请选择 PMTE 参数表 Excel 文件")
        self.output_dir_input = self.create_file_input("请选择输出目录")

        prior_btn = self.create_browse_btn("选择目录")
        prior_btn.clicked.connect(self.browse_prior_dir)
        pmte_btn = self.create_browse_btn("选择文件")
        pmte_btn.clicked.connect(self.browse_pmte)
        output_btn = self.create_browse_btn("选择目录")
        output_btn.clicked.connect(self.browse_output_dir)

        layout.addWidget(self.create_label("上年底稿目录"), 0, 0)
        layout.addWidget(self.prior_dir_input, 0, 1)
        layout.addWidget(prior_btn, 0, 2)
        layout.addWidget(self.create_label("PMTE 参数表"), 1, 0)
        layout.addWidget(self.pmte_input, 1, 1)
        layout.addWidget(pmte_btn, 1, 2)
        layout.addWidget(self.create_label("输出目录"), 2, 0)
        layout.addWidget(self.output_dir_input, 2, 1)
        layout.addWidget(output_btn, 2, 2)

        card.layout().addLayout(layout)
        return card

    def create_subject_card(self):
        card = self.create_card("科目选择")
        layout = QVBoxLayout()
        layout.setSpacing(12)

        self.subject_checkboxes = {}
        subject_grid = QGridLayout()
        subject_grid.setHorizontalSpacing(12)
        subject_grid.setVerticalSpacing(10)
        subject_grid.setColumnStretch(0, 1)
        subject_grid.setColumnStretch(1, 1)
        subject_grid.setColumnStretch(2, 1)

        try:
            config_manager = SubjectConfig()
            for index, (code, name) in enumerate(config_manager.get_subject_list()):
                checkbox = QCheckBox(f"{code}    {name}")
                checkbox.setObjectName("SubjectCheckbox")
                checkbox.setMinimumHeight(44)
                self.subject_checkboxes[code] = checkbox
                subject_grid.addWidget(checkbox, index // 3, index % 3)
        except Exception as exc:
            error_label = QLabel(f"加载失败: {exc}")
            error_label.setStyleSheet(f"color: {EY_ERROR};")
            subject_grid.addWidget(error_label, 0, 0, 1, 3)

        button_row = QHBoxLayout()
        select_all_btn = self.create_browse_btn("全选 / 取消")
        select_all_btn.setObjectName("SecondaryButton")
        select_all_btn.clicked.connect(self.toggle_select_all)
        button_row.addWidget(select_all_btn)
        button_row.addStretch()

        layout.addLayout(subject_grid)
        layout.addLayout(button_row)
        card.layout().addLayout(layout)
        return card

    def create_action_card(self):
        card = self.create_card("执行")
        layout = QVBoxLayout()
        layout.setSpacing(12)

        button_row = QHBoxLayout()
        button_row.setSpacing(12)
        self.start_btn = QPushButton("开始处理")
        self.start_btn.setObjectName("PrimaryButton")
        self.start_btn.clicked.connect(self.start_processing)
        self.clear_btn = QPushButton("一键清空")
        self.clear_btn.setObjectName("SecondaryButton")
        self.clear_btn.clicked.connect(self.clear_form)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)

        button_row.addWidget(self.start_btn, 3)
        button_row.addWidget(self.clear_btn, 1)
        layout.addLayout(button_row)
        layout.addWidget(self.progress_bar)
        card.layout().addLayout(layout)
        return card

    def create_log_card(self):
        card = self.create_card("处理日志")
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setPlaceholderText("等待处理...")
        self.log_output.setMaximumHeight(190)
        card.layout().addWidget(self.log_output)
        return card

    def create_results_card(self):
        card = self.create_card("处理结果")
        self.results_table = QTableWidget(0, 5)
        self.results_table.setHorizontalHeaderLabels(["科目", "状态", "输出文件", "Warnings", "Wording 匹配数量"])
        self.results_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.results_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.results_table.verticalHeader().setVisible(False)
        self.results_table.setMinimumHeight(170)
        self.results_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.results_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.results_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        card.layout().addWidget(self.results_table)
        return card

    def create_footer(self):
        footer = QWidget()
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(0, 4, 0, 0)
        left = QLabel("READY | Audit Roll Forward test build")
        left.setStyleSheet(f"color: {EY_MUTED}; font-size: 11px;")
        right = QLabel("EY black / yellow theme")
        right.setStyleSheet(f"color: {EY_MUTED}; font-size: 11px;")
        layout.addWidget(left)
        layout.addStretch()
        layout.addWidget(right)
        return footer

    def create_label(self, text):
        label = QLabel(text)
        label.setObjectName("FieldLabel")
        return label

    def create_input(self, placeholder):
        line_edit = QLineEdit()
        line_edit.setPlaceholderText(placeholder)
        return line_edit

    def create_file_input(self, placeholder):
        line_edit = QLineEdit()
        line_edit.setPlaceholderText(placeholder)
        line_edit.setReadOnly(False)
        return line_edit

    def create_browse_btn(self, text):
        return QPushButton(text)

    def clear_form(self):
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "提示", "当前正在处理，请等待完成后再清空。")
            return

        for field in (
            self.company_input,
            self.date_input,
            self.functional_currency_input,
            self.accounting_standard_input,
            self.pm_input,
            self.te_input,
            self.prior_dir_input,
            self.pmte_input,
            self.output_dir_input,
        ):
            field.clear()

        self.roll_wording_checkbox.setChecked(False)
        self.generate_summary_checkbox.setChecked(True)
        self.keep_adjustments_checkbox.setChecked(False)
        self.clear_testing_checkbox.setChecked(False)

        if hasattr(self, "subject_checkboxes"):
            for checkbox in self.subject_checkboxes.values():
                checkbox.setChecked(False)

        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)
        self.log_output.clear()
        self.results_table.setRowCount(0)

    def toggle_select_all(self):
        if hasattr(self, "subject_checkboxes"):
            should_check = not all(checkbox.isChecked() for checkbox in self.subject_checkboxes.values())
            for checkbox in self.subject_checkboxes.values():
                checkbox.setChecked(should_check)

    def log_event(self, message):
        try:
            APP_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(APP_LOG_PATH, "a", encoding="utf-8") as log_file:
                log_file.write(f"[{timestamp}] {message}\n")
        except Exception:
            pass

    def dialog_helper_command(self):
        if getattr(sys, "frozen", False):
            helper_path = Path(sys.executable).with_name("AuditRollForward_DialogHelper.exe")
            if not helper_path.exists():
                raise FileNotFoundError(f"找不到选择器组件: {helper_path}")
            return [str(helper_path)]

        helper_path = Path(__file__).with_name("dialog_helper.py")
        if not helper_path.exists():
            raise FileNotFoundError(f"找不到选择器脚本: {helper_path}")
        return [sys.executable, str(helper_path)]

    def run_external_dialog_helper(self, kind, title, result_path, filter_text=""):
        command = self.dialog_helper_command() + [
            kind,
            "--result",
            str(result_path),
            "--title",
            title,
        ]
        if filter_text:
            command.extend(["--filter", filter_text])
        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=creationflags,
        )
        if completed.returncode != 0:
            stderr = (completed.stderr or completed.stdout or "").strip()
            raise RuntimeError(stderr or f"Dialog helper exited with code {completed.returncode}")
        result_file = Path(result_path)
        if not result_file.exists():
            return ""
        return result_file.read_text(encoding="utf-8-sig").strip()

    def choose_directory_external(self, title):
        self.log_event(f"open_helper_directory_dialog: {title}")
        result_path = tempfile.NamedTemporaryFile(delete=False, suffix=".txt").name
        Path(result_path).unlink(missing_ok=True)
        try:
            path = self.run_external_dialog_helper("directory", title, result_path)
            self.log_event(f"helper_directory_selected: {path}")
            return path or ""
        except Exception as exc:
            self.log_event(f"helper_directory_error: {exc}")
            QMessageBox.warning(self, "提示", f"打开目录选择窗口失败：{exc}\n\n也可以直接把目录路径粘贴到输入框。")
            return ""
        finally:
            Path(result_path).unlink(missing_ok=True)

    def choose_file_external(self, title, filter_text):
        self.log_event(f"open_helper_file_dialog: {title}")
        result_path = tempfile.NamedTemporaryFile(delete=False, suffix=".txt").name
        Path(result_path).unlink(missing_ok=True)
        try:
            path = self.run_external_dialog_helper("file", title, result_path, filter_text)
            self.log_event(f"helper_file_selected: {path}")
            return path or ""
        except Exception as exc:
            self.log_event(f"helper_file_error: {exc}")
            QMessageBox.warning(self, "提示", f"打开文件选择窗口失败：{exc}\n\n也可以直接把文件路径粘贴到输入框。")
            return ""
        finally:
            Path(result_path).unlink(missing_ok=True)

    def browse_prior_dir(self):
        dir_path = self.choose_directory_external("选择上年底稿目录")
        if dir_path:
            self.prior_dir_input.setText(dir_path)

    def browse_pmte(self):
        file_path = self.choose_file_external(
            "选择 PMTE 参数表",
            "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*",
        )
        if file_path:
            self.pmte_input.setText(file_path)

    def browse_output_dir(self):
        dir_path = self.choose_directory_external("选择输出目录")
        if dir_path:
            self.output_dir_input.setText(dir_path)

    def feedback_prompt_already_shown(self):
        try:
            if not FEEDBACK_STATE_PATH.exists():
                return False
            with open(FEEDBACK_STATE_PATH, "r", encoding="utf-8") as state_file:
                state = json.load(state_file)
            return bool(state.get("feedback_prompt_shown"))
        except Exception:
            return False

    def save_feedback_prompt_state(self, opened=False):
        try:
            APP_STATE_DIR.mkdir(parents=True, exist_ok=True)
            state = {
                "feedback_prompt_shown": True,
                "feedback_opened": bool(opened),
                "feedback_prompt_shown_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "feedback_url": FEEDBACK_URL,
            }
            with open(FEEDBACK_STATE_PATH, "w", encoding="utf-8") as state_file:
                json.dump(state, state_file, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def maybe_show_feedback_prompt(self, success_count, total_count):
        if total_count <= 0 or success_count <= 0 or self.feedback_prompt_already_shown():
            return

        message = QMessageBox(self)
        message.setWindowTitle("使用反馈")
        message.setIcon(QMessageBox.Icon.Information)
        message.setText("本次 Roll Forward 已完成。")
        message.setInformativeText("为了继续优化工具，是否愿意花 1 分钟填写使用反馈？")
        fill_button = message.addButton("填写问卷", QMessageBox.ButtonRole.AcceptRole)
        later_button = message.addButton("暂不填写", QMessageBox.ButtonRole.RejectRole)
        message.setDefaultButton(fill_button)
        message.exec()

        opened = message.clickedButton() == fill_button
        self.save_feedback_prompt_state(opened=opened)
        if opened:
            QDesktopServices.openUrl(QUrl(FEEDBACK_URL))

    def start_processing(self):
        subject_codes = [
            code for code, checkbox in getattr(self, "subject_checkboxes", {}).items()
            if checkbox.isChecked()
        ]
        if not subject_codes:
            QMessageBox.warning(self, "提示", "请至少选择一个科目。")
            return

        prior_dir = self.prior_dir_input.text().strip()
        pmte_path = self.pmte_input.text().strip()
        company_name = self.company_input.text().strip()
        bs_date = self.date_input.text().strip()
        output_dir = self.output_dir_input.text().strip()

        if not all([prior_dir, pmte_path, company_name, bs_date, output_dir]):
            QMessageBox.warning(self, "提示", "请填写公司名称、资产负债表日，并选择上年底稿、PMTE 和输出目录。")
            return

        template_dir = resource_path("templates")
        if not os.path.exists(template_dir):
            QMessageBox.warning(self, "提示", f"找不到模板目录: {template_dir}")
            return

        self.start_btn.setEnabled(False)
        self.start_btn.setText("处理中...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(subject_codes))
        self.progress_bar.setValue(0)
        self.results_table.setRowCount(0)
        self.log_output.clear()

        self.log_output.append(">>> 初始化完成")
        self.log_output.append(f">>> 目标公司: {company_name}")
        self.log_output.append(f">>> 资产负债表日: {bs_date}")
        if self.functional_currency_input.text().strip():
            self.log_output.append(f">>> 记账本位币: {self.functional_currency_input.text().strip()}")
        if self.accounting_standard_input.text().strip():
            self.log_output.append(f">>> 适用会计准则: {self.accounting_standard_input.text().strip()}")
        if self.pm_input.text().strip():
            self.log_output.append(f">>> PM: {self.pm_input.text().strip()}")
        if self.te_input.text().strip():
            self.log_output.append(f">>> TE: {self.te_input.text().strip()}")
        self.log_output.append(f">>> 待处理科目数: {len(subject_codes)}")
        if self.roll_wording_checkbox.isChecked():
            self.log_output.append(">>> Wording roll forward 已启用")
        if self.generate_summary_checkbox.isChecked():
            self.log_output.append(">>> Roll Forward Summary 已启用")
        self.log_output.append(">>> " + "=" * 50)

        self.worker = RollForwardWorker(
            subject_codes=subject_codes,
            template_dir=template_dir,
            prior_dir=prior_dir,
            pmte_path=pmte_path,
            company_name=company_name,
            bs_date=bs_date,
            output_dir=output_dir,
            functional_currency=self.functional_currency_input.text().strip(),
            accounting_standard=self.accounting_standard_input.text().strip(),
            pm_value=self.pm_input.text().strip(),
            te_value=self.te_input.text().strip(),
            roll_forward_wording=self.roll_wording_checkbox.isChecked(),
            generate_summary=self.generate_summary_checkbox.isChecked(),
        )
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.processing_finished)
        self.worker.start()

    def update_progress(self, current, total, message):
        self.progress_bar.setMaximum(max(total, 1))
        self.progress_bar.setValue(current)
        self.log_output.append(message)

    def populate_results_table(self, results):
        self.results_table.setRowCount(len(results))
        for row, result in enumerate(results):
            subject_code, success, message, output_path, warnings_list = result[:5]
            metadata = getattr(warnings_list, "metadata", {}) if warnings_list is not None else {}
            warning_text = "; ".join(str(item) for item in warnings_list) if warnings_list else ""
            values = [
                subject_code,
                "成功" if success else "失败",
                output_path or "",
                warning_text or message,
                str(metadata.get("wording_copied_count", 0)),
            ]

            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col == 1:
                    item.setForeground(QColor(EY_SUCCESS if success else EY_ERROR))
                self.results_table.setItem(row, col, item)
        self.results_table.resizeRowsToContents()

    def processing_finished(self, results):
        self.start_btn.setEnabled(True)
        self.start_btn.setText("开始处理")

        success_count = sum(1 for result in results if len(result) > 1 and result[1])
        total_count = len(results)
        self.populate_results_table(results)
        self.log_output.append(">>> " + "=" * 50)
        self.log_output.append(f">>> 处理完成: {success_count}/{total_count}")

        if total_count and success_count == total_count:
            QMessageBox.information(self, "完成", f"全部科目处理成功，共生成 {success_count} 个底稿文件。")
        else:
            QMessageBox.warning(self, "完成", f"部分科目处理失败，成功 {success_count}/{total_count}。")

        self.maybe_show_feedback_prompt(success_count, total_count)


class RollForwardWorker(QThread):
    progress_signal = pyqtSignal(int, int, str)
    finished_signal = pyqtSignal(list)

    def __init__(
        self,
        subject_codes,
        template_dir,
        prior_dir,
        pmte_path,
        company_name,
        bs_date,
        output_dir,
        functional_currency=None,
        accounting_standard=None,
        pm_value=None,
        te_value=None,
        roll_forward_wording=False,
        generate_summary=True,
    ):
        super().__init__()
        self.subject_codes = subject_codes
        self.template_dir = template_dir
        self.prior_dir = prior_dir
        self.pmte_path = pmte_path
        self.company_name = company_name
        self.bs_date = bs_date
        self.output_dir = output_dir
        self.functional_currency = functional_currency
        self.accounting_standard = accounting_standard
        self.pm_value = pm_value
        self.te_value = te_value
        self.roll_forward_wording = roll_forward_wording
        self.generate_summary = generate_summary

    def run(self):
        try:
            def progress_callback(current, total, message):
                self.progress_signal.emit(current, total, message)

            results = process_multiple_subjects(
                self.subject_codes,
                self.template_dir,
                self.prior_dir,
                self.pmte_path,
                self.company_name,
                self.bs_date,
                self.output_dir,
                functional_currency=self.functional_currency,
                accounting_standard=self.accounting_standard,
                pm_value=self.pm_value,
                te_value=self.te_value,
                roll_forward_wording=self.roll_forward_wording,
                generate_summary=self.generate_summary,
                progress_callback=progress_callback,
            )

            self.finished_signal.emit(results)
        except Exception as exc:
            self.progress_signal.emit(0, 0, f"错误: {exc}")
            self.finished_signal.emit([])


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 10))
    window = RollForwardApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
