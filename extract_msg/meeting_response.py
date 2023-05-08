__all__ = [
    'MeetingResponse',
]


import datetime

from typing import Optional

from . import constants
from .enums import ResponseType
from .meeting_related import MeetingRelated


class MeetingResponse(MeetingRelated):
    """
    Class for handling meeting response objects.
    """

    @property
    def appointmentCounterProposal(self) -> bool:
        """
        Indicates if the response is a counter proposal.
        """
        return self._ensureSetNamed('_appointmentCounterProposal', '8257', constants.PSETID_APPOINTMENT, overrideClass = bool, preserveNone = False)

    @property
    def appointmentProposedDuration(self) -> Optional[int]:
        """
        The proposed value for the appointmentDuration property for a counter
        proposal.
        """
        return self._ensureSetNamed('_appointmentProposedDuration', '8256', constants.PSETID_APPOINTMENT)

    @property
    def appointmentProposedEndWhole(self) -> Optional[datetime.datetime]:
        """
        The proposal value for the appointmentEndWhole property for a counter
        proposal.
        """
        return self._ensureSetNamed('_appointmentProposedEndWhole', '8251', constants.PSETID_APPOINTMENT)

    @property
    def appointmentProposedStartWhole(self) -> Optional[datetime.datetime]:
        """
        The proposal value for the appointmentStartWhole property for a counter
        proposal.
        """
        return self._ensureSetNamed('_appointmentProposedStartWhole', '8250', constants.PSETID_APPOINTMENT)

    @property
    def isSilent(self) -> bool:
        """
        Indicates if the user did not include any text in the body of the
        Meeting Response object.
        """
        return self._ensureSetNamed('_isSilent', '0004', constants.PSETID_MEETING, overrideClass = bool, preserveNone = False)

    @property
    def promptSendUpdate(self) -> bool:
        """
        Indicates that the Meeting Response object was out-of-date when it was
        received.
        """
        return self._ensureSetNamed('_promptSendUpdate', '8045', constants.PSETID_COMMON, overrideClass = bool, preserveNone = False)

    @property
    def responseType(self) -> Optional[ResponseType]:
        """
        The type of Meeting Response object.
        """
        # The ending of the class type determines the type of response.
        return ResponseType(self.classType.lower().split('.')[-1])
