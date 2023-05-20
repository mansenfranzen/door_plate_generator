import warnings
from typing import Dict

import pandas as pd

from pptx_slide_generator.models import SlidesData, SlideData


def _sanitize_values(row: pd.Series) -> Dict:
    """Helper function to sanitize raw excel values.

    """

    row[row.isnull()] = ""
    return row.to_dict()


def load_slide_content(column_key: str,
                       column_template: str,
                       path: str) -> SlidesData:
    """Loads input slides data.

    """

    df = pd.read_excel(path)
    slides_data = {}
    for _, row in df.iterrows():
        key = row[column_key]
        layout = row[column_template]

        if pd.isnull(layout):
            warnings.warn(f"Row {key} has no layout.")
            layout = None

        slide_data = SlideData(key=key,
                               layout=layout,
                               values=_sanitize_values(row))

        slides_data[key] = slide_data

    return SlidesData(slides=slides_data)
