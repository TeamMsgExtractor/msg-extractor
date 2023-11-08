__all__ = [
    'Journal',
]


import base64
import datetime
import functools
import json

from typing import List, Optional

from ..constants import HEADER_FORMAT_TYPE, ps
from ..enums import LogFlags
from .message_base import MessageBase
from ..utils import minutesToDurationStr


class Journal(MessageBase):
    """
    Class for parsing Journal messages.
    """

    def getJson(self) -> str:
        return json.dumps({
            'subject': self.subject,
            'entryType': self.logTypeDesc,
            'company': self.companies[0] if self.companies else None,
            'start': self.logStart.__format__(self.datetimeFormat) if self.logStart else None,
            'end': self.logEnd.__format__(self.datetimeFormat) if self.logEnd else None,
            'duration': minutesToDurationStr(self.duration),
            'body': self.body,
            # There is a good chance the body property won't exist, so this is a
            # backup.
            'rtfBodyB64': base64.b64encode(self.rtfBody).decode('ascii') if self.rtfBody else None,
        })

    @functools.cached_property
    def companies(self) -> Optional[List[str]]:
        """
        Contains a list of company names, each of which is accociated with a
        contact this is precified in the contacts property.
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
    def logDuration(self) -> int:
        """
        The duration, in minutes, of the activity.
        """
        return self.getNamedProp('8707', ps.PSETID_LOG, 0)

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
        return {
            '-main details-': {
                'Subject': self.subject,
                'Entry Type': self.logTypeDesc,
                'Company': self.companies[0] if self.companies else None,
            },
            '-time-': {
                'Start': self.logStart.__format__(self.datetimeFormat) if self.logStart else None,
                'End': self.logEnd.__format__(self.datetimeFormat) if self.logEnd else None,
                'Duration': minutesToDurationStr(self.duration),
            },
        }
