import io

import pandas as pd


def read_excel(data: bytes) -> pd.DataFrame:
    """Read the first sheet as strings so cell values stay stable.

    Without ``dtype=str`` pandas turns numeric cells into floats (``123.0``)
    and dates into timestamps, which leak into generated filenames and Word
    content. Empty cells become ``""`` instead of ``NaN``.
    """
    df = pd.read_excel(
        io.BytesIO(data),
        sheet_name=0,
        dtype=str,
        keep_default_na=False,
    )
    df.columns = [str(c).strip() for c in df.columns]
    return df
