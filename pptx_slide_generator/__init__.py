from pptx_slide_generator.excel import load_slide_content
from pptx_slide_generator.powerpoint import rename_shapes, generate_slides
from pptx_slide_generator.svg import get_svg_shape_names

from pptx import Presentation


def run(excel_path: str,
        excel_column_key: str,
        excel_column_template: str,
        pptx_path: str,
        pptx_slide_idx: int,
        pptx_shape_exclude: str,
        pptx_shape_prefix: str,
        svg_path: str,
        svg_name_attribute: str,
        result_path: str):
    """Main entry point to execute generate slides from svg and excel input."""

    slides_data = load_slide_content(column_key=excel_column_key,
                                     column_template=excel_column_template,
                                     path=excel_path)

    svg_names = get_svg_shape_names(path=svg_path,
                                    id_attribute=svg_name_attribute)

    pptx = Presentation(pptx_path)

    rename_shapes(pptx=pptx,
                  slide_idx=pptx_slide_idx,
                  names=svg_names,
                  exclude=pptx_shape_exclude)

    generate_slides(pptx=pptx,
                    slide_idx=pptx_slide_idx,
                    slides_data=slides_data,
                    prefix=pptx_shape_prefix,
                    exclude=pptx_shape_exclude)

    pptx.save(result_path)
