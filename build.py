#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyInstaller 打包脚本
生成便携版 exe
"""

import PyInstaller.__main__
import os

# 打包参数
PyInstaller.__main__.run([
    'main_gui.py',
    '--name=AuditRollForward',
    '--onefile',
    '--windowed',
    '--noconfirm',
    '--icon=NONE',
    '--add-data=subjects_config.json;.',  # Windows下用分号，Linux/Mac下用冒号
    '--add-data=templates;templates',
    '--clean',
])
