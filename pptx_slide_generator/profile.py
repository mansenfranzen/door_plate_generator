import json
from pathlib import Path
from typing import Dict
from loguru import logger

PROFILE_FILENAME = "pptx_slide_generator.json"
PREVIOUS_PROFILE_NAME = "LATEST"


def add_profile(profile_name: str, values: Dict):
    """Stores given profile in home directory as json file. If profile
    doesn't exists, it will be created.

    """
    path = Path().home().joinpath(PROFILE_FILENAME)
    if path.exists():
        profiles = json.loads(path.read_text(encoding="utf-8"))
    else:
        profiles = {}

    parameters = {key: value for key, value in values.items() if value is not None}
    profiles[profile_name] = parameters
    path.write_text(json.dumps(profiles, indent=3), encoding="utf-8")

    logger.info(f"Store '{profile_name} under '{path}' with following values:\n\n{json.dumps(parameters, indent=3)}")


def add_previous_profile(values: Dict):
    """Store the latest profile under special name.

    """

    add_profile(PREVIOUS_PROFILE_NAME, values)


def add_default_profile():
    """Store default profile."""
    values = dict(excel_path="PathToExcelFile",
                  pptx_path="PathToPptxFile",
                  svg_path="PathToSvgFile",
                  result_path="PathToResultPptxFile")

    add_profile("default", values)


def read_profile(profile_name: str) -> Dict:
    """Load profile from profile store.

    """

    path = Path().home().joinpath(PROFILE_FILENAME)

    try:
        profiles = json.loads(path.read_text(encoding="utf-8"))
        return profiles[profile_name]

    except (KeyError, FileNotFoundError) as e:
        raise e("Profile not found. Please create first.")


def read_previous_profile() -> Dict:
    """Load last profile from profile store.

    """
    return read_profile(PREVIOUS_PROFILE_NAME)
