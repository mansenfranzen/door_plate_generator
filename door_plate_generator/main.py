from loguru import logger

from door_plate_generator.excel import load_excel_content
from door_plate_generator.powerpoint import rename_shapes, generate_slides
from door_plate_generator.svg import get_svg_shape_names
from pptx import Presentation


def run(**kwargs):
    """Main entry point to execute generate slides from svg and excel input."""

    logger_name = f'logs/{kwargs["excel_section_value"]}-{{time}}.log'
    logger.add(logger_name, level="WARNING", format="{message}")
    logger.info(f'Processing {kwargs["pptx_path"]}')

    excel_data = load_excel_content(column_room=kwargs["excel_column_room"],
                                    column_layout=kwargs["excel_column_layout"],
                                    column_section=kwargs["excel_column_section"],
                                    column_relevant=kwargs["excel_column_relevant"],
                                    section_value=kwargs["excel_section_value"],
                                    path=kwargs["excel_path"])

    svg_names = get_svg_shape_names(path=kwargs["svg_path"],
                                    id_attribute=kwargs["svg_name_attribute"])

    pptx = Presentation(kwargs["pptx_path"])

    rename_shapes(pptx=pptx,
                  slide_idx=kwargs["pptx_slide_idx"],
                  names=svg_names,
                  exclude=kwargs["pptx_shape_exclude"])

    generate_slides(pptx=pptx,
                    slide_idx=kwargs["pptx_slide_idx"],
                    excel_data=excel_data,
                    prefix=kwargs["pptx_shape_prefix"],
                    exclude=kwargs["pptx_shape_exclude"])

    pptx.save(kwargs["result_path"])
