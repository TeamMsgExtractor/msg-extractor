import datetime
import logging

from .enums import TaskAcceptance, TaskHistory, TaskMode, TaskOwnership, TaskStatus
from .message_base import MessageBase


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Task(MessageBase):
    """
    Class used for parsing task files.
    """

    def __init__(self, path, **kwargs):
        super().__init__(path, **kwargs)

    @property
    def percentComplete(self) -> float:
        """
        Indicates whether a time-flagged Message object is complete. Returns a
        percentage in decimal form. 1.0 indicates it is complete.
        """
        return self._ensureSetNamed('_percentComplete', '8102')

    @property
    def taskAcceptanceState(self) -> TaskAcceptance:
        """
        Indicates the acceptance state of the task.
        """
        return self._ensureSetNamed('_percentComplete', '812A', overrideClass = TaskAcceptance)

    @property
    def taskActualEffort(self) -> int:
        """
        Indicates the number of minutes that the user actually spent working on
        a task.
        """
        return self._ensureSetNamed('_taskActualEffort', '8110')

    @property
    def taskAssigner(self) -> str:
        """
        Specifies the name of the user that last assigned the task.
        """
        return self._ensureSetNamed('_taskAssigner', '811F')

    @property
    def taskComplete(self) -> bool:
        """
        Indicates if the task is complete.
        """
        return self._ensureSetNamed('_taskComplete', '811C')

    @property
    def taskCustomFlags(self) -> int:
        """
        Custom flags set on the task.
        """
        return self._ensureSetNamed('_taskCustomFlags', '8139')

    @property
    def taskDueDate(self) -> datetime.datetime:
        """
        Specifies the date by which the user expects work on the task to be
        complete.
        """
        return self._ensureSetNamed('_taskStartDate', '8105')

    @property
    def taskEstimatedEffort(self) -> int:
        """
        Indicates the number of minutes that the user expects to work on a task.
        """
        return self._ensureSetNamed('_taskEstimatedEffort', '8111')

    @property
    def taskFRecurring(self) -> bool:
        """
        Indicates whether the task includes a recurrence pattern.
        """
        return self._ensureSetNamed('_taskFRecurring', '8126')

    @property
    def taskHistory(self) -> TaskHistory:
        """
        Indicates the type of change that was last made to the Task object.
        """
        return self._ensureSetNamed('_taskHistory', '811A', overrideClass = TaskHistory)

    @property
    def taskLastDelegate(self) -> str:
        """
        Contains the name of the user who most recently assigned the task, or
        the user to whom it was most recently assigned.
        """
        return self._ensureSetNamed('_taskLastDelegate', '8125')

    @property
    def taskLastUser(self) -> str:
        """
        Contains the name of the most recent user to have been the owner of the
        task.
        """
        return self._ensureSetNamed('_taskLastUser', '8122')

    @property
    def taskMode(self) -> TaskMode:
        """
        Used in a task communication. Should be 0 (UNASSIGNED) on task objects.
        """
        return self._ensureSetNamed('_taskMode', '8518', overrideClass = TaskMode)

    @property
    def taskOwner(self) -> str:
        """
        Contains the name of the owner of the task.
        """
        return self._ensureSetNamed('_taskOwner', '811F')

    @property
    def taskOwnership(self) -> TaskOwnership:
        """
        Contains the name of the owner of the task.
        """
        return self._ensureSetNamed('_taskOwnership', '8129', overrideClass = TaskOwnership)

    @property
    def taskStartDate(self) -> datetime.datetime:
        """
        Specifies the date on which the user expects work on the task to begin.
        """
        return self._ensureSetNamed('_taskStartDate', '8104')

    @property
    def taskStatus(self) -> TaskStatus:
        """
        The completion status of a task.
        """
        return self._ensureSetNamed('_taskStatus', '8101', overrideClass = TaskStatus)

    @property
    def taskStatusOnComplete(self) -> bool:
        """
        Indicates whether the task assignee has been requested to send an email
        message upon completion of the assigned task.
        """
        return self._ensureSetNamed('_taskStatusOnComplete', '8119')

    @property
    def taskUpdates(self) -> bool:
        """
        Indicates whether the task assignee has been requested to send a task
        update when the assigned Task object changes.
        """
        return self._ensureSetNamed('_taskUpdates', '811B')

    @property
    def taskVersion(self) -> int:
        """
        Indicates which copy is the latest update of a Task object.
        """
        return self._ensureSetNamed('_taskVersion', '8113')
