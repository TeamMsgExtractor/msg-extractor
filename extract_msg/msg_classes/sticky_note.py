import functools
import json

from typing import Optional

from .. import constants
from ..enums import NoteColor
from .message_base import MessageBase


# Note: Sticky note is basically just text and a background color, so we don't
# really do much when saving it.
class StickyNote(MessageBase):
    """
    A sticky note.
    """

    def getJson(self) -> str:
        return json.dumps({
            'subject': self.subject,
            'date': self.date.__format__(self.datetimeFormat) if self.date else None,
            'body': self.body,
            'height': self.noteHeight,
            'width': self.noteWidth,
            'color': None if self.noteColor is None else self.noteColor.name.lower(),
        })

    @property
    def headerFormatProperties(self) -> constants.HEADER_FORMAT_TYPE:
        return None

    @functools.cached_property
    def noteColor(self) -> Optional[NoteColor]:
        """
        The color of the sticky note.
        """
        return self.getNamedAs('8B00', constants.ps.PSETID_NOTE, NoteColor)

    @functools.cached_property
    def noteHeight(self) -> Optional[int]:
        """
        The height of the note window, in pixels.
        """
        return self.getNamedProp('8B03', constants.ps.PSETID_NOTE)

    @functools.cached_property
    def noteWidth(self) -> Optional[int]:
        """
        The width of the note window, in pixels.
        """
        return self.getNamedProp('8B02', constants.ps.PSETID_NOTE)

    @functools.cached_property
    def noteX(self) -> Optional[int]:
        """
        The distance, in pixels, from the left edge of the screen that a user
        interface displays the note.
        """
        return self.getNamedProp('8B02', constants.ps.PSETID_NOTE)

    @functools.cached_property
    def noteY(self) -> Optional[int]:
        """
        The distance, in pixels, from the top edge of the screen that a user
        interafce displays the note.
        """
        return self.getNamedProp('8B02', constants.ps.PSETID_NOTE)