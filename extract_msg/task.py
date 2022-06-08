import datetime
import logging

from .enums import TaskMode, TaskStatus
from .msg import MSGFile

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

class Task(MSGFile):
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
    def taskDueDate(self) -> datetime.datetime:
        """
        Specifies the date by which the user expects work on the task to be
        complete.
        """
        return self._ensureSetNamed('_taskStartDate', '8105')

    @property
    def taskMode(self) -> TaskMode:
        """
        Used in a task communication. Should be 0 (UNASSIGNED) on task objects.
        """
        return self._ensureSetNamed('_taskMode', '8518', overrideClass = TaskMode)

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
