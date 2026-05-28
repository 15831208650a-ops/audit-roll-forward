#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
审计底稿 Roll Forward 桌面应用 (v4.0 - 赛博朋克风格)
技术栈：PyQt6 + openpyxl + PyInstaller
作者：AI Assistant
"""

import sys
import os

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QFileDialog, QCheckBox,
                             QProgressBar, QTextEdit, QGroupBox, QMessageBox, QListWidget,
                             QListWidgetItem, QFrame, QScrollArea, QGridLayout, QGraphicsDropShadowEffect)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QFont, QColor, QPalette, QLinearGradient, QBrush, QPainter, QPaintEvent

from roll_forward_core import process_multiple_subjects, resource_path, SubjectConfig


class CyberpunkApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AUDIT SYSTEM // ROLL FORWARD v4.0")
        self.setGeometry(100, 80, 1200, 900)
        self.setStyleSheet(self.get_cyberpunk_stylesheet())

        self.worker = None
        self.init_ui()

    def get_cyberpunk_stylesheet(self):
        """赛博朋克主样式表"""
        return """
            /* ========== 全局基础 ========== */
            QWidget {
                font-family: "Microsoft YaHei", "Segoe UI", sans-serif;
                font-size: 13px;
                color: #e0e0e0;
            }

            QMainWindow {
                background-color: #0a0a0f;
            }

            /* ========== 滚动区域 ========== */
            QScrollArea {
                background-color: transparent;
                border: none;
            }

            QScrollBar:vertical {
                background-color: #0a0a0f;
                width: 6px;
                border: none;
            }

            QScrollBar::handle:vertical {
                background: #1a1d2e;
                min-height: 30px;
            }

            QScrollBar::handle:vertical:hover {
                background: #00f0ff;
            }

            /* ========== 霓虹输入框 ========== */
            QLineEdit {
                background-color: rgba(10, 10, 15, 0.9);
                border: 1px solid #1a1d2e;
                color: #ffffff;
                padding: 10px 14px;
                font-size: 13px;
                selection-background-color: #00f0ff;
                selection-color: #000000;
            }

            QLineEdit:focus {
                border: 1px solid #00f0ff;
            }

            QLineEdit::placeholder {
                color: #444;
            }

            /* ========== 文件路径框 ========== */
            QLineEdit#file_path {
                background-color: rgba(10, 10, 15, 0.9);
                border: 1px solid #1a1d2e;
                color: #888;
                padding: 10px 14px;
            }

            QLineEdit#file_path:focus {
                border: 1px solid #00f0ff;
                color: #fff;
            }

            /* ========== 浏览按钮 ========== */
            QPushButton#browse_btn {
                background-color: transparent;
                color: #00f0ff;
                border: 1px solid rgba(0, 240, 255, 0.4);
                padding: 8px 16px;
                font-weight: bold;
                font-size: 12px;
            }

            QPushButton#browse_btn:hover {
                background-color: rgba(0, 240, 255, 0.08);
                border: 1px solid #00f0ff;
            }

            /* ========== 霓虹主按钮 ========== */
            QPushButton#neon_btn {
                background-color: rgba(0, 240, 255, 0.08);
                color: #00f0ff;
                border: 1px solid #00f0ff;
                padding: 14px 32px;
                font-weight: bold;
                font-size: 13px;
                letter-spacing: 2px;
            }

            QPushButton#neon_btn:hover {
                background-color: rgba(0, 240, 255, 0.15);
            }

            /* ========== 洋红按钮 ========== */
            QPushButton#magenta_btn {
                background-color: transparent;
                color: #ff00a0;
                border: 1px solid #ff00a0;
                padding: 8px 16px;
                font-weight: bold;
                font-size: 12px;
            }

            QPushButton#magenta_btn:hover {
                background-color: rgba(255, 0, 160, 0.08);
            }

            /* ========== 科目列表 ========== */
            QListWidget {
                background-color: rgba(10, 10, 15, 0.8);
                border: 1px solid #1a1d2e;
                padding: 8px;
                outline: none;
            }

            QListWidget::item {
                padding: 10px 14px;
                margin: 4px;
                border: 1px solid transparent;
                color: #888;
            }

            QListWidget::item:selected {
                background-color: rgba(0, 240, 255, 0.08);
                border-left: 3px solid #00f0ff;
                color: #00f0ff;
            }

            QListWidget::item:hover {
                background-color: rgba(0, 240, 255, 0.04);
            }

            /* ========== 进度条 ========== */
            QProgressBar {
                border: none;
                background-color: #1a1d2e;
                text-align: center;
                color: #fff;
                font-weight: bold;
                height: 24px;
            }

            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #00f0ff,
                    stop:1 #ff00a0);
            }

            /* ========== 日志输出 ========== */
            QTextEdit {
                background-color: rgba(10, 10, 15, 0.9);
                border: 1px solid #1a1d2e;
                padding: 12px;
                color: #aaa;
                font-family: "Consolas", "Courier New", monospace;
                font-size: 12px;
                line-height: 1.8;
            }

            /* ========== 复选框 ========== */
            QCheckBox {
                spacing: 8px;
                font-size: 13px;
                color: #e0e0e0;
            }

            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border: 1px solid #1a1d2e;
                background-color: transparent;
            }

            QCheckBox::indicator:checked {
                background-color: #00f0ff;
                border: 1px solid #00f0ff;
            }
        """

    def init_ui(self):
        """初始化赛博朋克界面"""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("background-color: transparent; border: none;")
        self.setCentralWidget(scroll)

        container = QWidget()
        scroll.setWidget(container)
        layout = QVBoxLayout(container)
        layout.setSpacing(24)
        layout.setContentsMargins(40, 32, 40, 32)

        # ========== 顶部标题区 ==========
        header = self.create_header()
        layout.addWidget(header)

        # ========== 基础参数卡片 ==========
        params_card = self.create_glass_card("基础参数", "▼")
        params_layout = QGridLayout()
        params_layout.setSpacing(16)
        params_layout.setColumnStretch(1, 3)
        params_layout.setColumnStretch(3, 3)

        # 公司名称
        params_layout.addWidget(self.create_label("公司名称"), 0, 0)
        self.company_input = self.create_cyber_input("请输入公司名称")
        params_layout.addWidget(self.company_input, 0, 1)

        # 资产负债表日期
        params_layout.addWidget(self.create_label("资产负债表日期"), 0, 2)
        self.date_input = self.create_cyber_input("格式: 2026/12/31")
        params_layout.addWidget(self.date_input, 0, 3)

        # 记账本位币
        params_layout.addWidget(self.create_label("记账本位币"), 1, 0)
        self.functional_currency_input = self.create_cyber_input("如：人民币")
        params_layout.addWidget(self.functional_currency_input, 1, 1)

        # 适用会计准则
        params_layout.addWidget(self.create_label("适用会计准则"), 1, 2)
        self.accounting_standard_input = self.create_cyber_input("如：企业会计准则")
        params_layout.addWidget(self.accounting_standard_input, 1, 3)

        params_card.layout().addLayout(params_layout)
        layout.addWidget(params_card)

        # ========== 文件路径卡片 ==========
        file_card = self.create_glass_card("文件路径", "▼")
        file_layout = QVBoxLayout()
        file_layout.setSpacing(12)

        # 上年底稿目录
        prior_row = QHBoxLayout()
        prior_row.addWidget(self.create_label("上年底稿目录"), 1)
        self.prior_dir_input = self.create_file_input("点击选择目录...")
        prior_row.addWidget(self.prior_dir_input, 5)
        prior_browse_btn = self.create_browse_btn()
        prior_browse_btn.clicked.connect(self.browse_prior_dir)
        prior_row.addWidget(prior_browse_btn, 1)
        file_layout.addLayout(prior_row)

        # PMTE信息表
        pmte_row = QHBoxLayout()
        pmte_row.addWidget(self.create_label("PMTE信息表"), 1)
        self.pmte_input = self.create_file_input("点击选择文件...")
        pmte_row.addWidget(self.pmte_input, 5)
        pmte_browse_btn = self.create_browse_btn()
        pmte_browse_btn.clicked.connect(self.browse_pmte)
        pmte_row.addWidget(pmte_browse_btn, 1)
        file_layout.addLayout(pmte_row)

        # 输出目录
        output_row = QHBoxLayout()
        output_row.addWidget(self.create_label("输出目录"), 1)
        self.output_dir_input = self.create_file_input("点击选择目录...")
        output_row.addWidget(self.output_dir_input, 5)
        output_browse_btn = self.create_browse_btn()
        output_browse_btn.clicked.connect(self.browse_output_dir)
        output_row.addWidget(output_browse_btn, 1)
        file_layout.addLayout(output_row)

        file_card.layout().addLayout(file_layout)
        layout.addWidget(file_card)

        # ========== 科目选择卡片 ==========
        subject_card = self.create_glass_card("科目选择", "▼")
        subject_layout = QVBoxLayout()
        subject_layout.setSpacing(12)

        # 科目列表
        self.subject_list = QListWidget()
        self.subject_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.subject_list.setMaximumHeight(220)
        self.subject_list.setStyleSheet("""
            QListWidget {
                background-color: rgba(10, 10, 15, 0.8);
                border: 1px solid #1a1d2e;
                padding: 8px;
            }
            QListWidget::item {
                padding: 10px 14px;
                margin: 4px;
                border: 1px solid transparent;
                color: #888;
            }
            QListWidget::item:selected {
                background-color: rgba(0, 240, 255, 0.08);
                border-left: 3px solid #00f0ff;
                color: #00f0ff;
            }
        """)

        # 加载科目
        try:
            config_manager = SubjectConfig()
            subjects = config_manager.get_subject_list()
            for code, name in subjects:
                item = QListWidgetItem(f"  {code}    {name}")
                item.setData(Qt.ItemDataRole.UserRole, code)
                self.subject_list.addItem(item)
        except Exception as e:
            self.subject_list.addItem(f"加载失败: {str(e)}")

        subject_layout.addWidget(self.subject_list)

        # 全选按钮
        btn_row = QHBoxLayout()
        select_all_btn = self.create_magenta_btn("全选 / 取消")
        select_all_btn.clicked.connect(self.toggle_select_all)
        btn_row.addWidget(select_all_btn)
        btn_row.addStretch()
        subject_layout.addLayout(btn_row)

        subject_card.layout().addLayout(subject_layout)
        layout.addWidget(subject_card)

        # ========== 操作区域 ==========
        action_card = self.create_glass_card("操作", "▼")
        action_layout = QVBoxLayout()
        action_layout.setSpacing(16)

        # 开始按钮
        self.start_btn = self.create_neon_btn("▶  开始处理")
        self.start_btn.clicked.connect(self.start_processing)
        action_layout.addWidget(self.start_btn)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        action_layout.addWidget(self.progress_bar)

        action_card.layout().addLayout(action_layout)
        layout.addWidget(action_card)

        # ========== 日志区域 ==========
        log_card = self.create_glass_card("系统日志", "▼")
        log_layout = QVBoxLayout()

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setPlaceholderText(">>> 等待指令...")
        self.log_output.setMaximumHeight(200)
        log_layout.addWidget(self.log_output)

        log_card.layout().addLayout(log_layout)
        layout.addWidget(log_card)

        # ========== 底部状态栏 ==========
        footer = self.create_footer()
        layout.addWidget(footer)

    def create_header(self):
        """创建赛博朋克风格标题"""
        header = QWidget()
        header_layout = QVBoxLayout(header)
        header_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_layout.setSpacing(8)

        # Logo图标
        logo_row = QHBoxLayout()
        logo_row.setAlignment(Qt.AlignmentFlag.AlignCenter)

        logo_icon = QLabel("◈")
        logo_icon.setStyleSheet("""
            color: #00f0ff;
            font-size: 32px;
            margin-right: 12px;
        """)
        logo_row.addWidget(logo_icon)

        # 主标题
        title_label = QLabel("AUDIT SYSTEM")
        title_label.setStyleSheet("""
            color: #ffffff;
            font-size: 28px;
            font-weight: bold;
            letter-spacing: 4px;
        """)
        logo_row.addWidget(title_label)
        header_layout.addLayout(logo_row)

        # 副标题
        subtitle = QLabel("ROLL FORWARD PROTOCOL v4.0")
        subtitle.setStyleSheet("""
            color: #00f0ff;
            font-size: 11px;
            letter-spacing: 6px;
            margin-top: 4px;
        """)
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_layout.addWidget(subtitle)

        # 装饰线
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("""
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 transparent,
                stop:0.3 #00f0ff,
                stop:0.7 #ff00a0,
                stop:1 transparent);
            min-height: 1px;
            max-height: 1px;
            margin-top: 16px;
        """)
        header_layout.addWidget(line)

        return header

    def create_glass_card(self, title, icon=""):
        """创建玻璃态卡片"""
        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: rgba(15, 17, 25, 0.8);
                border: 1px solid #1a1d2e;
                padding: 20px;
            }
        """)
        layout = QVBoxLayout(card)
        layout.setSpacing(16)

        # 标题
        title_label = QLabel(f"{icon}  {title}")
        title_label.setStyleSheet("""
            color: #ffffff;
            font-size: 14px;
            font-weight: bold;
            letter-spacing: 2px;
            padding-bottom: 8px;
            border-bottom: 1px solid #1a1d2e;
        """)
        layout.addWidget(title_label)

        return card

    def create_label(self, text):
        """创建标签"""
        label = QLabel(text)
        label.setStyleSheet("""
            color: #00f0ff;
            font-size: 11px;
            letter-spacing: 2px;
            text-transform: uppercase;
        """)
        return label

    def create_cyber_input(self, placeholder):
        """创建赛博朋克输入框"""
        input_field = QLineEdit()
        input_field.setPlaceholderText(placeholder)
        input_field.setStyleSheet("""
            QLineEdit {
                background-color: rgba(10, 10, 15, 0.9);
                border: 1px solid #1a1d2e;
                color: #ffffff;
                padding: 10px 14px;
            }
            QLineEdit:focus {
                border: 1px solid #00f0ff;
            }
        """)
        return input_field

    def create_file_input(self, placeholder):
        """创建文件路径输入框"""
        input_field = QLineEdit()
        input_field.setPlaceholderText(placeholder)
        input_field.setObjectName("file_path")
        input_field.setReadOnly(True)
        input_field.setStyleSheet("""
            QLineEdit {
                background-color: rgba(10, 10, 15, 0.9);
                border: 1px solid #1a1d2e;
                color: #888;
                padding: 10px 14px;
            }
        """)
        return input_field

    def create_browse_btn(self):
        """创建浏览按钮"""
        btn = QPushButton("浏览")
        btn.setObjectName("browse_btn")
        return btn

    def create_neon_btn(self, text):
        """创建霓虹主按钮"""
        btn = QPushButton(text)
        btn.setObjectName("neon_btn")
        return btn

    def create_magenta_btn(self, text):
        """创建洋红色按钮"""
        btn = QPushButton(text)
        btn.setObjectName("magenta_btn")
        return btn

    def create_footer(self):
        """创建底部状态栏"""
        footer = QWidget()
        footer.setStyleSheet("padding-top: 16px;")
        footer_layout = QHBoxLayout(footer)

        left = QLabel("READY // AUDIT ROLL FORWARD SYSTEM v4.0")
        left.setStyleSheet("color: #555; font-size: 10px; letter-spacing: 2px;")
        footer_layout.addWidget(left)

        footer_layout.addStretch()

        right = QLabel("CPU: 12%    MEM: 456MB")
        right.setStyleSheet("color: #555; font-size: 10px; letter-spacing: 1px;")
        right.setAlignment(Qt.AlignmentFlag.AlignRight)
        footer_layout.addWidget(right)

        return footer

    def toggle_select_all(self):
        """全选/取消全选"""
        if self.subject_list.selectedItems():
            self.subject_list.clearSelection()
        else:
            self.subject_list.selectAll()

    def browse_prior_dir(self):
        """浏览上年底稿目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择上年底稿目录")
        if dir_path:
            self.prior_dir_input.setText(dir_path)
            self.prior_dir_input.setStyleSheet("""
                QLineEdit {
                    background-color: rgba(10, 10, 15, 0.9);
                    border: 1px solid #00f0ff;
                    color: #fff;
                    padding: 10px 14px;
                }
            """)

    def browse_pmte(self):
        """浏览PMTE信息表"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择PMTE信息表", "", "Excel文件 (*.xlsx)")
        if file_path:
            self.pmte_input.setText(file_path)
            self.pmte_input.setStyleSheet("""
                QLineEdit {
                    background-color: rgba(10, 10, 15, 0.9);
                    border: 1px solid #00f0ff;
                    color: #fff;
                    padding: 10px 14px;
                }
            """)

    def browse_output_dir(self):
        """浏览输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_dir_input.setText(dir_path)
            self.output_dir_input.setStyleSheet("""
                QLineEdit {
                    background-color: rgba(10, 10, 15, 0.9);
                    border: 1px solid #00f0ff;
                    color: #fff;
                    padding: 10px 14px;
                }
            """)

    def start_processing(self):
        """开始处理"""
        selected_items = self.subject_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请至少选择一个科目！")
            return

        subject_codes = [item.data(Qt.ItemDataRole.UserRole) for item in selected_items]

        prior_dir = self.prior_dir_input.text().strip()
        pmte_path = self.pmte_input.text().strip()
        company_name = self.company_input.text().strip()
        bs_date = self.date_input.text().strip()
        functional_currency = self.functional_currency_input.text().strip()
        accounting_standard = self.accounting_standard_input.text().strip()
        output_dir = self.output_dir_input.text().strip()

        if not all([prior_dir, pmte_path, company_name, bs_date, output_dir]):
            QMessageBox.warning(self, "警告", "请填写所有必填项！")
            return

        template_dir = resource_path("templates")
        if not os.path.exists(template_dir):
            QMessageBox.warning(self, "警告", f"找不到模板目录: {template_dir}")
            return

        self.start_btn.setEnabled(False)
        self.start_btn.setText("⏳ 处理中...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(subject_codes))
        self.progress_bar.setValue(0)
        self.log_output.clear()

        self.log_output.append(">>> 系统初始化完成")
        self.log_output.append(f">>> 目标公司: {company_name}")
        self.log_output.append(f">>> 资产负债表日: {bs_date}")
        if functional_currency:
            self.log_output.append(f">>> 记账本位币: {functional_currency}")
        if accounting_standard:
            self.log_output.append(f">>> 适用会计准则: {accounting_standard}")
        self.log_output.append(f">>> 待处理科目数: {len(subject_codes)}")
        self.log_output.append(">>> " + "=" * 50)

        self.worker = RollForwardWorker(
            subject_codes, template_dir, prior_dir, pmte_path,
            company_name, bs_date, output_dir,
            functional_currency, accounting_standard
        )
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.processing_finished)
        self.worker.start()

    def update_progress(self, current, total, message):
        """更新进度"""
        self.progress_bar.setValue(current)
        self.log_output.append(message)

    def processing_finished(self, results):
        """处理完成"""
        self.start_btn.setEnabled(True)
        self.start_btn.setText("▶  开始处理")

        success_count = sum(1 for result in results if len(result) > 1 and result[1])
        total_count = len(results)

        self.log_output.append(">>> " + "=" * 50)
        self.log_output.append(f">>> 处理完成！成功: {success_count}/{total_count}")
        self.log_output.append(">>> 系统就绪")

        if success_count == total_count:
            QMessageBox.information(self, "完成", f"所有科目处理成功！\n共生成 {success_count} 个底稿文件。")
        else:
            QMessageBox.warning(self, "完成", f"部分科目处理失败！\n成功: {success_count}/{total_count}")


class RollForwardWorker(QThread):
    """后台处理线程"""
    progress_signal = pyqtSignal(int, int, str)
    finished_signal = pyqtSignal(list)

    def __init__(self, subject_codes, template_dir, prior_dir, pmte_path,
                 company_name, bs_date, output_dir,
                 functional_currency=None, accounting_standard=None):
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

    def run(self):
        try:
            results = process_multiple_subjects(
                self.subject_codes, self.template_dir, self.prior_dir,
                self.pmte_path, self.company_name, self.bs_date, self.output_dir,
                functional_currency=self.functional_currency,
                accounting_standard=self.accounting_standard
            )

            for i, (subject_code, success, message, output_path, warnings_list) in enumerate(results):
                status = "成功" if success else "失败"
                self.progress_signal.emit(i + 1, len(results), f"[{status}] {subject_code}: {message}")

            self.finished_signal.emit(results)

        except Exception as e:
            self.progress_signal.emit(0, 0, f"错误: {str(e)}")
            self.finished_signal.emit([])


def main():
    app = QApplication(sys.argv)

    # 设置全局字体
    font = QFont("Microsoft YaHei", 10)
    app.setFont(font)

    window = CyberpunkApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
