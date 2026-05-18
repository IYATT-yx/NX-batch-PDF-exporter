"""
file: language.py
description: 语言模块
author: IYATT-yx
repository: https://github.com/IYATT-yx/NX-batch-PDF-exporter
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
"""
import os
import builtins
import gettext
import ctypes

def get_windows_language_string() -> str:
    """从 Windows 系统读取语言字符串
    """
    try:
        len = 85
        buf = ctypes.create_unicode_buffer(len)
        ctypes.windll.kernel32.GetUserDefaultLocaleName(buf, len)
        win_lang = buf.value
        if win_lang:
            return win_lang.replace('-', '_')
    except Exception:
        pass
    return "zh_CN"

def init_i18n(domain="messages", locales_dir=None) -> str:
    """初始化语言国际化

    Returns:
        str: 当前语言
    """
    if locales_dir is None:
        locales_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'locales')
    current_lang = get_windows_language_string()
    translation = gettext.translation(
        domain=domain,
        localedir=locales_dir,
        languages=[current_lang],
        fallback=True
    )
    builtins._ = translation.gettext
    builtins.n_ = translation.ngettext
    return current_lang