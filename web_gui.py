#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
审计底稿 Roll Forward - 网页版GUI (v4.1)
使用 QWebEngineView + QWebChannel 实现网页前端 + Python 后端
"""

import sys
import os
import json

from PyQt6.QtWidgets import (QApplication, QMainWindow, QFileDialog, QMessageBox)
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtCore import Qt, QUrl, pyqtSlot, QObject, QThread, pyqtSignal
from PyQt6.QtGui import QFont
from PyQt6.QtWebChannel import QWebChannel

from roll_forward_core import process_multiple_subjects, SubjectConfig


class RollForwardWorker(QThread):
    """后台处理线程"""
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool)

    def __init__(self, subject_codes, template_dir, prior_dir, pmte_path,
                 company_name, bs_date, output_dir):
        super().__init__()
        self.subject_codes = subject_codes
        self.template_dir = template_dir
        self.prior_dir = prior_dir
        self.pmte_path = pmte_path
        self.company_name = company_name
        self.bs_date = bs_date
        self.output_dir = output_dir

    def run(self):
        try:
            results = process_multiple_subjects(
                self.subject_codes, self.template_dir, self.prior_dir,
                self.pmte_path, self.company_name, self.bs_date, self.output_dir
            )

            for i, (subject_code, success, message, output_path, warnings_list) in enumerate(results):
                status = "成功" if success else "失败"
                self.log_signal.emit(f"[{status}] {subject_code}: {message}")

            success_count = sum(1 for _, success, _, _, _ in results if success)
            self.log_signal.emit(f">>> 处理完成！成功: {success_count}/{len(results)}")
            self.finished_signal.emit(success_count == len(results))

        except Exception as e:
            self.log_signal.emit(f">>> 错误: {str(e)}")
            self.finished_signal.emit(False)


class WebBridge(QObject):
    """Python-JavaScript 桥接对象"""
    log_message = pyqtSignal(str)
    processing_finished = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.worker = None
        self._window = None

    def set_window(self, window):
        self._window = window

    @pyqtSlot(str, result=str)
    def browseDirectory(self, title="选择目录"):
        """浏览目录 - 返回选择的目录路径"""
        if self._window:
            dir_path = QFileDialog.getExistingDirectory(self._window, title)
            return dir_path if dir_path else ""
        return ""

    @pyqtSlot(str, result=str)
    def browseFile(self, title="选择文件"):
        """浏览文件 - 返回选择的文件路径"""
        if self._window:
            file_path, _ = QFileDialog.getOpenFileName(
                self._window, title, "", "Excel文件 (*.xlsx)"
            )
            return file_path if file_path else ""
        return ""

    @pyqtSlot(str, result=list)
    def getSubjects(self, _=""):
        """获取科目列表"""
        try:
            config_manager = SubjectConfig()
            subjects = config_manager.get_subject_list()
            return subjects
        except Exception as e:
            return []

    @pyqtSlot(str, str, str, str, str, str, result=bool)
    def startProcessing(self, subject_codes_json, prior_dir, pmte_path,
                        company_name, bs_date, output_dir):
        """开始处理"""
        try:
            subject_codes = json.loads(subject_codes_json)
            template_dir = os.path.join(os.path.dirname(__file__), "templates")

            if not os.path.exists(template_dir):
                self.log_message.emit(f">>> 错误: 找不到模板目录 {template_dir}")
                return False

            self.log_message.emit(f">>> 开始处理 {len(subject_codes)} 个科目...")

            self.worker = RollForwardWorker(
                subject_codes, template_dir, prior_dir, pmte_path,
                company_name, bs_date, output_dir
            )
            self.worker.log_signal.connect(self.on_worker_log)
            self.worker.finished_signal.connect(self.on_worker_finished)
            self.worker.start()
            return True

        except Exception as e:
            self.log_message.emit(f">>> 错误: {str(e)}")
            return False

    def on_worker_log(self, message):
        """接收工作线程日志"""
        self.log_message.emit(message)

    def on_worker_finished(self, success):
        """处理完成"""
        self.processing_finished.emit(success)


class CyberpunkWebApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AUDIT SYSTEM // ROLL FORWARD v4.1")
        self.setGeometry(100, 80, 1400, 900)
        self.setStyleSheet("background-color: #0a0a0f;")

        # 创建 WebView
        self.web_view = QWebEngineView()
        self.setCentralWidget(self.web_view)

        # 创建桥接对象
        self.bridge = WebBridge(self)
        self.bridge.set_window(self)

        # 设置 QWebChannel
        self.channel = QWebChannel()
        self.channel.registerObject("pyqtObj", self.bridge)
        self.web_view.page().setWebChannel(self.channel)

        # 加载本地HTML文件
        html_path = os.path.join(os.path.dirname(__file__), "cyberpunk_ui.html")
        self.web_view.setUrl(QUrl.fromLocalFile(html_path))

    def closeEvent(self, event):
        """关闭窗口时停止工作线程"""
        if self.bridge.worker and self.bridge.worker.isRunning():
            self.bridge.worker.terminate()
            self.bridge.worker.wait(1000)
        event.accept()


def main():
    app = QApplication(sys.argv)

    # 设置全局字体
    font = QFont("Microsoft YaHei", 10)
    app.setFont(font)

    window = CyberpunkWebApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
