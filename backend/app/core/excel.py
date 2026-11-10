import io

import pandas as pd

from app.core.config import settings


def list_sheets(data: bytes) -> list[str]:
    workbook = pd.ExcelFile(io.BytesIO(data))
    return [str(name) for name in workbook.sheet_names]


def read_excel(data: bytes, sheet_name: str | int = 0) -> pd.DataFrame:
    """Read one sheet as strings so cell values stay stable.

    Without ``dtype=str`` pandas turns numeric cells into floats (``123.0``)
    and dates into timestamps, which leak into generated filenames and Word
    content. Empty cells become ``""`` instead of ``NaN``.
    """
    df = pd.read_excel(
        io.BytesIO(data),
        sheet_name=sheet_name,
        dtype=str,
        keep_default_na=False,
    )
    df.columns = [str(column).strip() for column in df.columns]

    if len(df) > settings.max_excel_rows:
        raise ValueError(f"Excel 数据行数超过限制（最多 {settings.max_excel_rows} 行）")
    if df.shape[0] * df.shape[1] > settings.max_excel_cells:
        raise ValueError(f"Excel 单元格总数超过限制（最多 {settings.max_excel_cells} 个）")
    return df
