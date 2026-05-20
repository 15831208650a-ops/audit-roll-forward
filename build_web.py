#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyInstaller 打包脚本 - 网页版GUI
生成便携版 exe
"""

import PyInstaller.__main__
import os

# 网页版打包参数
PyInstaller.__main__.run([
    'web_gui.py',
    '--name=AuditRollForward_Web',
    '--onefile',
    '--windowed',
    '--icon=NONE',
    '--add-data=subjects_config.json;.',
    '--add-data=templates;templates',
    '--add-data=cyberpunk_ui.html;.',
    '--add-data=qwebchannel.js;.',
    '--hidden-import=PyQt6.QtWebEngineCore',
    '--hidden-import=PyQt6.QtWebChannel',
    '--clean',
])
