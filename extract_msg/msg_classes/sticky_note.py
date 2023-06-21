from typing import Optional

from ..constants import HEADER_FORMAT_TYPE
from ..constants.ps import PSETID_NOTE
from ..enums import NoteColor
from .message_base import MessageBase


# Note: Sticky note is basically just text and a background color, so we don't
# really do much when saving it.
class StickyNote(MessageBase):
    """
    A sticky note.
    """

    @property
    def headerFormatProperties(self) -> HEADER_FORMAT_TYPE:
        return None

    @property
    def noteColor(self) -> Optional[NoteColor]:
        """
        The color of the sticky note.
        """
        return self._ensureSetNamed('_noteColor', '8B00', PSETID_NOTE, preserveNone = True, overrideClass = NoteColor)

    @property
    def noteHeight(self) -> Optional[int]:
        """
        The height of the note window, in pixels.
        """
        return self._ensureSetNamed('_noteWidth', '8B03', PSETID_NOTE)

    @property
    def noteWidth(self) -> Optional[int]:
        """
        The width of the note window, in pixels.
        """
        return self._ensureSetNamed('_noteWidth', '8B02', PSETID_NOTE)

    @property
    def noteX(self) -> Optional[int]:
        """
        The distance, in pixels, from the left edge of the screen that a user
        interface displays the note.
        """
        return self._ensureSetNamed('_noteX', '8B02', PSETID_NOTE)

    @property
    def noteY(self) -> Optional[int]:
        """
        The distance, in pixels, from the top edge of the screen that a user
        interafce displays the note.
        """
        return self._ensureSetNamed('_noteY', '8B02', PSETID_NOTE)