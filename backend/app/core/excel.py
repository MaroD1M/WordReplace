import io
import logging

import pandas as pd

from app.core.config import settings

logger = logging.getLogger("wordreplace")


def list_sheets(data: bytes) -> list[str]:
    workbook = pd.ExcelFile(io.BytesIO(data))
    return [str(name) for name in workbook.sheet_names]


def read_excel(data: bytes, sheet_name: str | int = 0) -> pd.DataFrame:
    """Read one sheet as strings so cell values stay stable.

    Without ``dtype=str`` pandas turns numeric cells into floats (``123.0``)
    and dates into timestamps, which leak into generated filenames and Word
    content. Empty cells become ``""`` instead of ``NaN``.

    If a requested sheet no longer exists (for example the user switched the
    Excel file while the page still held the previous sheet name), fall back to
    the first sheet instead of failing the whole preview/execute request.
    """
    workbook = pd.ExcelFile(io.BytesIO(data))
    sheets = [str(name) for name in workbook.sheet_names]

    target: str | int = sheet_name
    if isinstance(sheet_name, str):
        requested = sheet_name.strip()
        if requested and requested in sheets:
            target = requested
        else:
            target = sheets[0] if sheets else 0
            if requested:
                logger.warning(
                    "sheet %r not found, falling back to %r",
                    requested,
                    target,
                )

    df = pd.read_excel(
        workbook,
        sheet_name=target,
        dtype=str,
        keep_default_na=False,
    )
    df.columns = [str(column).strip() for column in df.columns]

    if len(df) > settings.max_excel_rows:
        raise ValueError(f"Excel 数据行数超过限制（最多 {settings.max_excel_rows} 行）")
    if df.shape[0] * df.shape[1] > settings.max_excel_cells:
        raise ValueError(f"Excel 单元格总数超过限制（最多 {settings.max_excel_cells} 个）")
    return df
