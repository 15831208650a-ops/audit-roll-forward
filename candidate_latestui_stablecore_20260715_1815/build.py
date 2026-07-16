#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyInstaller 打包脚本
生成便携版 exe
"""

import os
import sys

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
BUILD_DEPS = os.path.join(ROOT_DIR, ".build_deps")
if os.path.exists(BUILD_DEPS):
    sys.path.insert(0, BUILD_DEPS)

import PyInstaller.__main__

SPEC_DIR = os.path.join("build", "specs")
os.makedirs(SPEC_DIR, exist_ok=True)
ICON_PATH = os.path.join(ROOT_DIR, "assets", "audit_roll_forward_icon.ico")
CONFIG_PATH = os.path.join(ROOT_DIR, "subjects_config.json")
TEMPLATES_DIR = os.path.join(ROOT_DIR, "templates")
ASSETS_DIR = os.path.join(ROOT_DIR, "assets")

# 打包独立选择器组件。它不加载 PyQt/Tk/PowerShell，用于隔离系统文件选择窗口崩溃。
PyInstaller.__main__.run([
    'dialog_helper.py',
    '--name=AuditRollForward_DialogHelper',
    f'--specpath={SPEC_DIR}',
    '--onefile',
    '--windowed',
    '--noconfirm',
    f'--icon={ICON_PATH}',
    '--clean',
])

# 打包主程序
PyInstaller.__main__.run([
    'main_gui.py',
    '--name=AuditRollForward',
    f'--specpath={SPEC_DIR}',
    '--onefile',
    '--windowed',
    '--noconfirm',
    f'--icon={ICON_PATH}',
    f'--add-data={CONFIG_PATH};.',  # Windows下用分号，Linux/Mac下用冒号
    f'--add-data={TEMPLATES_DIR};templates',
    f'--add-data={ASSETS_DIR};assets',
    '--clean',
])
