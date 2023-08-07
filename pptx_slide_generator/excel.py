import functools
import warnings
from typing import Dict

import pandas as pd

from pptx_slide_generator.models import ExcelData, RoomData, SlideData


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


def _get_slide_data(row: pd.Series,
                    column_template: str,
                    column_relevant: str,
                    ) -> SlideData:
    """Extracts room data information from a single row in master
    excel file.

    """

    layout = row[column_template]
    relevant = row[column_relevant] == "ja"

    if pd.isnull(layout):
        layout = None

    values = _sanitize_values(row)

    return SlideData(layout=layout,
                     relevant=relevant,
                     values=values)


def load_excel_content(column_room: str,
                       column_layout: str,
                       column_section: str,
                       column_relevant: str,
                       section_value: str,
                       path: str) -> ExcelData:
    """Loads input slides data.

    """

    get_slide_data = functools.partial(
        _get_slide_data,
        column_template=column_layout,
        column_relevant=column_relevant
    )

    df = pd.read_excel(path, skiprows=2)
    df = _filter_rows(df=df,
                      column_section=column_section,
                      section_name=section_value)

    excel_data = {}
    df_grouped = df.groupby(column_room)
    for room_name, df_room in df_grouped:
        slides = [get_slide_data(row) for _, row in df_room.iterrows()]
        room_data = RoomData(name=room_name, slides=slides)
        excel_data[room_name] = room_data

    return ExcelData(rooms=excel_data)
