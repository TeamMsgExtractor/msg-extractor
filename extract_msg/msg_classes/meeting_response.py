__all__ = [
    'MeetingResponse',
]


import datetime
import functools

from typing import Optional

from ..constants import ps
from ..enums import ResponseType
from .meeting_related import MeetingRelated


class MeetingResponse(MeetingRelated):
    """
    Class for handling meeting response objects.
    """

    @functools.cached_property
    def appointmentCounterProposal(self) -> bool:
        """
        Indicates if the response is a counter proposal.
        """
        return bool(self.getNamedProp('8257', ps.PSETID_APPOINTMENT))

    @functools.cached_property
    def appointmentProposedDuration(self) -> Optional[int]:
        """
        The proposed value for the appointmentDuration property for a counter
        proposal.
        """
        return self.getNamedProp('8256', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def appointmentProposedEndWhole(self) -> Optional[datetime.datetime]:
        """
        The proposal value for the appointmentEndWhole property for a counter
        proposal.
        """
        return self.getNamedProp('8251', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def appointmentProposedStartWhole(self) -> Optional[datetime.datetime]:
        """
        The proposal value for the appointmentStartWhole property for a counter
        proposal.
        """
        return self.getNamedProp('8250', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def isSilent(self) -> bool:
        """
        Indicates if the user did not include any text in the body of the
        Meeting Response object.
        """
        return bool(self.getNamedProp('0004', ps.PSETID_MEETING))

    @functools.cached_property
    def promptSendUpdate(self) -> bool:
        """
        Indicates that the Meeting Response object was out-of-date when it was
        received.
        """
        return bool(self.getNamedProp('8045', ps.PSETID_COMMON))

    @functools.cached_property
    def responseType(self) -> ResponseType:
        """
        The type of Meeting Response object.
        """
        # The ending of the class type determines the type of response.
        return ResponseType(self.classType.lower().split('.')[-1])
