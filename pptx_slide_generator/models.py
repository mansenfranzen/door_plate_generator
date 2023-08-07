from typing import Dict, Optional, List

from pydantic import BaseModel


class SlideData(BaseModel):
    """Represents the data for an individual slide.

    """

    layout: Optional[str]
    relevant: bool
    values: Dict[str, str]


class RoomData(BaseModel):
    """Represents the data for an individual room.
    A single room may have one or more distinct slides.

    """

    name: str
    slides: List[SlideData]


class ExcelData(BaseModel):
    """Contains the data of all slides.

    """

    rooms: Dict[str, RoomData]

    def get(self, room_name: str) -> RoomData:
        return self.rooms[room_name]
