from typing import Tuple
import xml.etree.ElementTree as ET


def _get_xml_root_from_svg_path(path):
    return ET.parse(path).getroot()


def _get_g_root_elements(root):
    return root.findall("{http://www.w3.org/2000/svg}g")


def _has_children(element):
    return len(list(element.iter())) > 1


def get_svg_shape_names(path: str, id_attribute: str) -> Tuple[str]:
    """Loads svg file and gets names for top level groups.

    """

    root = _get_xml_root_from_svg_path(path)
    g_elements = _get_g_root_elements(root)
    ordered_names = [
        g_element.get(id_attribute)
        for g_element in g_elements
        if _has_children(g_element)
    ]

    return tuple(ordered_names)
