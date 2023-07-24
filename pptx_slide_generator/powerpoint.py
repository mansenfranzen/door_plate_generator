import sys
from typing import Dict, Tuple, List, Optional

from loguru import logger

from pptx.presentation import Presentation
from pptx.shapes.base import BaseShape
from pptx.slide import Slide

from pptx_slide_generator.models import SlidesData, SlideData


#logger.configure(handlers=[{"sink": sys.stderr, "format": "{message}", "level": "WARNING"}])
logger.add("logs/file_{time}.log", level="WARNING", format="{message}")

def _get_relevant_shapes_by_name(
        slide: Slide,
        exclude_name: Optional[str] = None,
        include_prefix: Optional[str] = None) -> List[BaseShape]:
    """Fetches all shapes for given slide which are not ignored.

    """

    shapes = slide.shapes

    if exclude_name:
        shapes = [shape for shape in shapes
                  if shape.name != exclude_name]

    if include_prefix:
        shapes = [shape for shape in shapes
                  if shape.name.startswith(include_prefix)]

    return shapes


def _get_sanitized_room_from_shape_name(shape: BaseShape, prefix: str):
    """Get the lookup key between pptx shape and excel row while sanitizing.

    """

    return shape.name.replace(prefix, "").replace("_", "/")


def get_named_master_layouts(pptx: Presentation) -> Dict:
    """Get master layouts with names as unique identifiers.

    """

    slide_layouts = {}

    for slide_master in pptx.slide_masters:
        for slide_layout in slide_master.slide_layouts:
            assert slide_layout.name not in slide_layouts
            slide_layouts[slide_layout.name] = slide_layout

    return slide_layouts


def rename_shapes(pptx: Presentation,
                  slide_idx: int,
                  names: Tuple[str],
                  exclude: str):
    """Renames shapes for a given slide index while ignoring shapes with a
    specific name.

    """

    slide = pptx.slides[slide_idx]
    shapes = _get_relevant_shapes_by_name(slide, exclude_name=exclude)
    assert len(shapes) == len(names)

    shape_names = zip(shapes, names)
    for shape, name in shape_names:
        shape.name = name


def _get_shape_coordinates(shape: BaseShape) -> Tuple[float, float]:
    """Get top and left coordinates of shape."""

    return shape.top, shape.left


def _mirror_shape_names(source: Slide, target: Slide):
    """Mirrors all shape names in source slide with target slice. Leverages
    shape coordinates to create mapping between source and target shape because
    order can't be assumed to be equal.

    """

    keyed_source_names = {_get_shape_coordinates(shape): shape.name
                          for shape in source.shapes}

    for shape in target.shapes:
        target_key = _get_shape_coordinates(shape)
        shape.name = keyed_source_names[target_key]


def _get_room_data(slides_data: SlidesData, room: str) -> Optional[SlideData]:
    """Retrieve slide data for given a room while performing validations.

    """

    try:
        data = slides_data.get(room)
    except KeyError:
        logger.warning(f"Room '{room}' from SVG/Pptx does not have any data in excel file.")
        return

    if not data.relevant:
        logger.info(f"Room '{room}' from SVG/Pptx is skipped because marked as not relevant in excel file.")
        return

    if not data.layout:
        logger.warning(f"Room '{room}' from SVG/Pptx is skipped because no layout provided in excel file.")
        return

    return data


def generate_slides(pptx: Presentation,
                    slide_idx: int,
                    slides_data: SlidesData,
                    prefix: str,
                    exclude: str):
    """Generates new slides given a powerpoint slide with relevant
    shape names and corresponding slide data.

    """

    slide = pptx.slides[slide_idx]
    shapes = _get_relevant_shapes_by_name(slide, include_prefix=prefix)
    shapes = sorted(shapes, key=lambda x: x.name)
    layouts = get_named_master_layouts(pptx)

    for shape in shapes:
        room = _get_sanitized_room_from_shape_name(shape, prefix)
        data = _get_room_data(slides_data=slides_data, room=room)

        if not data:
            continue

        try:
            layout = layouts[data.layout]
        except KeyError:
            logger.warning(f"Room '{room}' has no corresponding layout '{data.layout}' in Powerpoint master.")
            continue

        new_slide = pptx.slides.add_slide(layout)
        shape.click_action.target_slide = new_slide

        _populate_shapes(layout=layout,
                         new_slide=new_slide,
                         data=data,
                         ignore=exclude,
                         room=room)


def _populate_shapes(layout: Slide,
                     new_slide: Slide,
                     data: SlideData,
                     ignore: str,
                     room: str):
    """Populate newly created shape"""

    _mirror_shape_names(layout, new_slide)
    shapes = _get_relevant_shapes_by_name(new_slide, exclude_name=ignore)

    for shape in shapes:
        try:
            shape.text = data.values[shape.name]
        except KeyError:
            logger.warning(f"Room '{room}' with layout '{layout.name}' has no data for shape '{shape.name}' in excel.")
