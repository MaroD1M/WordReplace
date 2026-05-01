"""Core pure utilities for filename/text safety and preflight checks."""

import os
import re
import unicodedata
from datetime import datetime
from typing import List, Optional, Tuple, Set

import pandas as pd

MAX_FILENAME_LENGTH = 200


def clean_text(text: str) -> str:
    """清理文本：去除首尾空白、隐藏字符、特殊空格，统一格式。"""
    if not isinstance(text, str):
        return ""
    text = text.strip()
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[\u00A0\u2002-\u200B]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def clean_filename(filename: str) -> str:
    """清理文件名中的非法字符并规避路径穿越/保留设备名。"""
    name = unicodedata.normalize("NFKC", str(filename))
    name = re.sub(r"[\\/:*?\"<>|\x00-\x1f]", "_", name)
    name = name.replace("..", "_")
    name = name.strip().strip(".")
    if not name:
        name = "未命名"

    stem, ext = os.path.splitext(name)
    reserved = {
        "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
    }
    if stem.upper() in reserved:
        stem = f"_{stem}"
    return f"{stem}{ext}"


def sanitize_cache_key(filename: str) -> str:
    """限制缓存键名，防止路径注入。"""
    cleaned = re.sub(r"[^A-Za-z0-9_\-]", "_", str(filename))
    return cleaned[:120] or f"rule_{datetime.now().strftime('%Y%m%d_%H%M%S')}"


def generate_safe_filename(
    excel_row: pd.Series,
    file_name_col: str,
    file_prefix: str = "",
    file_suffix: str = "",
    row_idx: int = 0,
    max_length: int = MAX_FILENAME_LENGTH,
) -> str:
    """安全生成文件名，处理超长名称和特殊字符。"""
    try:
        if file_name_col and file_name_col in excel_row.index:
            base_name = clean_text(str(excel_row[file_name_col]))
        else:
            base_name = f"文件_{row_idx + 1}"

        if not base_name or base_name.isspace():
            base_name = f"文件_{row_idx + 1}"

        if file_prefix and file_suffix:
            filename = f"{file_prefix}{base_name}{file_suffix}.docx"
        elif file_prefix:
            filename = f"{file_prefix}{base_name}.docx"
        elif file_suffix:
            filename = f"{base_name}{file_suffix}.docx"
        else:
            filename = f"{base_name}.docx"

        filename = clean_filename(filename)

        if len(filename.encode("utf-8")) > max_length:
            truncated_base = base_name
            while len(f"{file_prefix}{truncated_base}{file_suffix}.docx".encode("utf-8")) > max_length and truncated_base:
                truncated_base = truncated_base[:-1]

            if file_prefix and file_suffix:
                filename = f"{file_prefix}{truncated_base}{file_suffix}.docx"
            elif file_prefix:
                filename = f"{file_prefix}{truncated_base}.docx"
            elif file_suffix:
                filename = f"{truncated_base}{file_suffix}.docx"
            else:
                filename = f"{truncated_base}.docx"

            filename = clean_filename(filename)

        return filename
    except Exception:
        return f"文件_{row_idx + 1}.docx"


def get_replace_blockers(
    word_file,
    excel_df: Optional[pd.DataFrame],
    replace_rules: List[Tuple[str, str]],
    start_row: int,
    end_row: int,
) -> List[str]:
    """返回替换前置条件不足的原因，用于UI禁用态提示。"""
    blockers = []
    if not word_file:
        blockers.append("请先上传Word模板")
    if excel_df is None or len(excel_df) == 0:
        blockers.append("请先上传非空Excel数据")
    if len(replace_rules) == 0:
        blockers.append("请至少添加1条替换规则")
    if start_row > end_row:
        blockers.append("起始行不能大于结束行")
    return blockers


def dedupe_filename(filename: str, used_names: Set[str]) -> str:
    """同批次生成文件名去重，避免下载时覆盖。"""
    safe_name = clean_filename(filename)
    if safe_name not in used_names:
        used_names.add(safe_name)
        return safe_name

    base, ext = os.path.splitext(safe_name)
    idx = 1
    candidate = f"{base}_{idx}{ext}"
    while candidate in used_names:
        idx += 1
        candidate = f"{base}_{idx}{ext}"

    used_names.add(candidate)
    return candidate
