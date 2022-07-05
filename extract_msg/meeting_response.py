import datetime

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
        return self._ensureSetNamed('_appointmentCounterProposal', '8257')

    @property
    def appointmentProposedDuration(self) -> int:
        """
        The proposed value for the appointmentDuration property for a counter
        proposal.
        """
        return self._ensureSetNamed('_appointmentProposedDuration', '8256')

    @property
    def appointmentProposedEndWhole(self) -> datetime.datetime:
        """
        The proposal value for the appointmentEndWhole property for a counter
        proposal.
        """
        return self._ensureSetNamed('_appointmentProposedEndWhole', '8251')

    @property
    def appointmentProposedStartWhole(self) -> datetime.datetime:
        """
        The proposal value for the appointmentStartWhole property for a counter
        proposal.
        """
        return self._ensureSetNamed('_appointmentProposedStartWhole', '8250')

    @property
    def isSilent(self) -> bool:
        """
        Indicates if the user did not include any text in the body of the
        Meeting Response object.
        """
        return self._ensureSetNamed('_isSilent', '0004')

    @property
    def promptSendUpdate(self) -> bool:
        """
        Indicates that the Meeting Response object was out-of-date when it was
        received.
        """
        return self._ensureSetNamed('_promptSendUpdate', '8045')

    @property
    def responseType(self) -> ResponseType:
        """
        The type of Meeting Response object.
        """
        # The ending of the class type determines the type of response.
        return ResponseType(self.classType.lower().split('.')[-1])
