# src/utils/helpers.py
"""通用工具函数"""
from pypinyin import lazy_pinyin, Style


def pinyin_key(name: str) -> str:
    """返回姓名的拼音字符串（小写），用于排序。"""
    if not name:
        return ""
    return "".join(lazy_pinyin(str(name))).lower()