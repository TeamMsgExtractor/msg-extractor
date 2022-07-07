from . import constants
from .enums import TaskMode, TaskRequestType
from .message_base import MessageBase


class TaskRequest(MessageBase):
    """

    """

    @property
    def headerFormatProperties(self) -> constants.HEADER_FORMAT_TYPE:
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
        """
        return {
            '-basic info-': {
                'Subject': self.subject,
            },
            '-status-': {
                'Status': None,
                'Percent Complete': None,
                'Date Completed': None,
            },
            '-work-': {
                'Total Work': None,
                'Actual Work': None,
            },
            '-owner-': {
                'Owner': None,
            },
            '-importance-': {
                'Importance': self.importanceString,
            },
        }

    @property
    def processed(self) -> bool:
        """
        Indicates whether a client has already processed a received task
        communication.
        """
        return self._ensureSetProperty('_taskMode', '7D01000B')

    @property
    def taskMode(self) -> TaskMode:
        """
        The assignment status of the embedded Task object.
        """
        return self._ensureSetNamed('_taskMode', '8518', constants.PSETID_COMMON, overrideClass = TaskMode)

    @property
    def taskRequestType(self) -> TaskRequestType:
        """
        The type of task request.
        """
        return self._ensureSet('_taskRequestType', '__substg1.0_001A', TaskRequestType.fromClassType)
