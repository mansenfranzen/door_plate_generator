import warnings
from typing import Dict

import pandas as pd

from pptx_slide_generator.models import SlidesData, SlideData


def _sanitize_values(row: pd.Series) -> Dict:
    """Helper function to sanitize raw excel values.

    """

    row[row.isnull()] = ""
    return row.to_dict()


def _filter_rows(df: pd.DataFrame,
                 column_section: str,
                 section_name: str,
                 ) -> pd.DataFrame:
    """Manages excel specific row filtering.

    """

    # remove first row which is only for human readability
    df = df.drop([0])

    # apply row filters
    mask = df[column_section].eq(section_name)

    df = df[mask]

    return df

def load_slide_content(column_key: str,
                       column_template: str,
                       column_section: str,
                       column_relevant: str,
                       relevant_name: str,
                       section_name: str,
                       path: str) -> SlidesData:
    """Loads input slides data.

    """

    df = pd.read_excel(path, skiprows=2)
    df = _filter_rows(df=df,
                      column_section=column_section,
                      section_name=section_name)

    slides_data = {}
    for _, row in df.iterrows():
        key = row[column_key]
        layout = row[column_template]
        relevant = row[column_relevant] == relevant_name

        if pd.isnull(layout):
            layout = None

        slide_data = SlideData(key=key,
                               layout=layout,
                               relevant=relevant,
                               values=_sanitize_values(row))

        slides_data[key] = slide_data

    return SlidesData(slides=slides_data)
