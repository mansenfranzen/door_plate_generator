from typing import Dict, Optional

from pydantic import BaseModel


class SlideData(BaseModel):
    """Represents the content for a single slide.

    """

    key: str
    layout: Optional[str]
    values: Dict[str, str]


class SlidesData(BaseModel):
    """Contains the data of all slides.

    """

    slides: Dict[str, SlideData]

    def get(self, slide_key: str) -> SlideData:
        return self.slides[slide_key]