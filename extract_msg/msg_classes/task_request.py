__all__ = [
    'TaskRequest',
]


import functools
import logging

from typing import Optional

from .. import constants
from ..enums import ErrorBehavior, TaskMode, TaskRequestType
from ..exceptions import StandardViolationError
from .message_base import MessageBase
from .task import Task


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class TaskRequest(MessageBase):
    """
    Class for handling Task Request objects, including Task Accept, Task
    Decline, and Task Update.
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

        If you class should not do *any* header injection, return None from this
        property.
        """
        # So this is rather weird. Looks like TaskRequest does not rely on
        # headers at all, simply using the body itself for all the data to
        # print. So I guess we just return None and handle that.
        return None

    @functools.cached_property
    def processed(self) -> bool:
        """
        Indicates whether a client has already processed a received task
        communication.
        """
        return self._getPropertyAs('7D01000B', bool, False)

    @functools.cached_property
    def taskMode(self) -> Optional[TaskMode]:
        """
        The assignment status of the embedded Task object.
        """
        return self._getNamedAs('8518', constants.ps.PSETID_COMMON, TaskMode)

    @functools.cached_property
    def taskObject(self) -> Optional[Task]:
        """
        The task object embedded in this Task Request object.

        This function does all of the most basic validation, and so will log
        most issues or throw exceptions if there are too many problems.

        :raises StandardViolationError: A standard was blatently violated in a
            way that program does not tolerate.
        """
        # Get the task object.
        #
        # The task object MUST be the first attachment, but we will be
        # lenient and allow it to be in any position. It not existing,
        # however, will not be tolerated.
        task = next(((index, att) for index, att in self.attachments if isinstance(att.data, Task)), None)

        if task is None:
            if ErrorBehavior.STANDARDS_VIOLATION in self.errorBehavior:
                logger.error('Task object not found on TaskRequest object.')
                return
            raise StandardViolationError('Task object not found on TaskRequest object.')

        # We know we have the task, let's make sure it's at index 0. If not,
        # log it.
        if task[0] != 0:
            logger.warning('Embedded task object was not located at index 0.')

        self._taskObject = task[1]

        return self._taskObject

    @functools.cached_property
    def taskRequestType(self) -> TaskRequestType:
        """
        The type of task request.
        """
        return self._getStreamAs('__substg1.0_001A', TaskRequestType.fromClassType)
