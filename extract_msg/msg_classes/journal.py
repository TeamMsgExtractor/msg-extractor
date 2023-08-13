__all__ = [
    'Journal',
]


import datetime
import functools

from typing import List, Optional

from ..constants import HEADER_FORMAT_TYPE, ps
from ..enums import LogFlags
from .message_base import MessageBase


class Journal(MessageBase):
    """
    Class for parsing Journal messages.
    """

    @functools.cached_property
    def companies(self) -> Optional[List[str]]:
        """
        The start time for the object.
        """
        return self.getNamedProp('8539', ps.PSETID_COMMON)

    @functools.cached_property
    def logDocumentPosted(self) -> bool:
        """
        Indicates whether the document was sent by email of posted to a server
        folder during journaling.
        """
        return bool(self.getNamedProp('8711', ps.PSETID_LOG))

    @functools.cached_property
    def logDocumentPrinted(self) -> bool:
        """
        Indicates whether the document was printed during journaling.
        """
        return bool(self.getNamedProp('870E', ps.PSETID_LOG))

    @functools.cached_property
    def logDocumentRouted(self) -> bool:
        """
        Indicates whether the document was sent to a routing recipient during
        journaling.
        """
        return bool(self.getNamedProp('8710', ps.PSETID_LOG))

    @functools.cached_property
    def logDocumentSaved(self) -> bool:
        """
        Indicates whether the document was saved during journaling.
        """
        return bool(self.getNamedProp('870F', ps.PSETID_LOG))

    @functools.cached_property
    def logDuration(self) -> Optional[int]:
        """
        The duration, in minutes, of the activity.
        """
        return self.getNamedProp('8707', ps.PSETID_LOG)

    @functools.cached_property
    def logEnd(self) -> Optional[datetime.datetime]:
        """
        The name of the activity that is being recorded.
        """
        return self.getNamedProp('8708', ps.PSETID_LOG)

    @functools.cached_property
    def logFlags(self) -> LogFlags:
        """
        The name of the activity that is being recorded.
        """
        return LogFlags(self.getNamedProp('870C', ps.PSETID_LOG, 0))

    @functools.cached_property
    def logStart(self) -> Optional[datetime.datetime]:
        """
        The name of the activity that is being recorded.
        """
        return self.getNamedProp('8706', ps.PSETID_LOG)

    @functools.cached_property
    def logType(self) -> Optional[str]:
        """
        The name of the activity that is being recorded.
        """
        return self.getNamedProp('8700', ps.PSETID_LOG)

    @functools.cached_property
    def logTypeDesc(self) -> Optional[str]:
        """
        The description of the activity that is being recorded.
        """
        return self.getNamedProp('8712', ps.PSETID_LOG)

    @property
    def headerFormatProperties(self) -> HEADER_FORMAT_TYPE:
        """
        Returns a dictionary of properties, in order, to be formatted into the
        header. Keys are the names to use in the header while the values are one
        of the following:
        None: Signifies no data was found for the property and it should be
            omitted from the header.
        str: A string to be formatted into the header using the string encoding.
        Tuple[Union[str, None], bool]: A string should be formatted into the
            header. If the bool is True, then place an empty string if the value
            is None, otherwise follow the same behavior as regular None.

        Additional note: If the value is an empty string, it will be dropped as
        well by default.

        Additionally you can group members of a header together by placing them
        in an embedded dictionary. Groups will be spaced out using a second
        instance of the join string. If any member of a group is being printed,
        it will be spaced apart from the next group/item.

        If you class should not do *any* header injection, return None from this
        property.
        """
        return {
            '-main details-': {
                'From': self.sender,
                'Posted At': self.date,
                'Conversation': self.conversation,
            },
            '-subject-': {
                'Subject': self.subject,
            },
            '-importance-': {
                'Importance': self.importanceString,
            },
        }
