__all__ = [
    'Task',
]


import datetime
import logging

from typing import Optional, Set

from . import constants
from .enums import (
        TaskAcceptance, TaskHistory, TaskMode, TaskMultipleRecipients,
        TaskOwnership, TaskState, TaskStatus
    )
from .message_base import MessageBase
from .structures.recurrence_pattern import RecurrencePattern
from .utils import unsignedToSignedInt


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Task(MessageBase):
    """
    Class used for parsing task files.
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
        status = {
            TaskStatus.NOT_STARTED: 'Not Started',
            TaskStatus.IN_PROGRESS: 'In Progress',
            TaskStatus.COMPLETE: 'Completed',
            TaskStatus.WAITING_ON_OTHER: 'Waiting on someone else',
            TaskStatus.DEFERRED: 'Deferred',
            None: None,
        }[self.taskStatus]

        return {
            '-basic info-': {
                'Subject': self.subject,
            },
            '-status-': {
                'Status': status,
                'Percent Complete': f'{self.percentComplete*100:.0f}%',
                'Date Completed': self.taskDateCompleted.__format__('%w, %B %d, %Y') if self.taskDateCompleted else None,
            },
            '-work-': {
                'Total Work': f'{self.taskEstimatedEffort or 0} minutes',
                'Actual Work': f'{self.taskActualEffort or 0} minutes',
            },
            '-owner-': {
                'Owner': self.taskOwner,
            },
            '-importance-': {
                'Importance': self.importanceString,
            },
        }

    @property
    def percentComplete(self) -> Optional[float]:
        """
        Indicates whether a time-flagged Message object is complete. Returns a
        percentage in decimal form. 1.0 indicates it is complete.
        """
        return self._ensureSetNamed('_percentComplete', '8102', constants.PSETID_TASK)

    @property
    def taskAcceptanceState(self) -> Optional[TaskAcceptance]:
        """
        Indicates the acceptance state of the task.
        """
        return self._ensureSetNamed('_taskAcceptanceState', '812A', constants.PSETID_TASK, overrideClass = TaskAcceptance)

    @property
    def taskAccepted(self) -> bool:
        """
        Indicates whether a task assignee has replied to a tesk request for this
        task object. Does not indicate if it was accepted or rejected.
        """
        return self._ensureSetNamed('_taskAccepted', '8108', constants.PSETID_TASK, overrideClass = bool, preserveNone = False)

    @property
    def taskActualEffort(self) -> Optional[int]:
        """
        Indicates the number of minutes that the user actually spent working on
        a task.
        """
        return self._ensureSetNamed('_taskActualEffort', '8110', constants.PSETID_TASK)

    @property
    def taskAssigner(self) -> Optional[str]:
        """
        Specifies the name of the user that last assigned the task.
        """
        return self._ensureSetNamed('_taskAssigner', '8121', constants.PSETID_TASK)

    @property
    def taskAssigners(self) -> Optional[bytes]:
        """
        A stack of entries, each representing a task assigner. The most recent
        task assigner (that is, the top) appears at the bottom.

        The documentation on this is weird, so I don't know how to parse it.
        """
        return self._ensureSetNamed('_taskAssigners', '8117', constants.PSETID_TASK)

    @property
    def taskComplete(self) -> bool:
        """
        Indicates if the task is complete.
        """
        return self._ensureSetNamed('_taskComplete', '811C', constants.PSETID_TASK, overrideClass = bool, preserveNone = False)

    @property
    def taskCustomFlags(self) -> Optional[int]:
        """
        Custom flags set on the task.
        """
        return self._ensureSetNamed('_taskCustomFlags', '8139', constants.PSETID_TASK)

    @property
    def taskDateCompleted(self) -> Optional[datetime.datetime]:
        """
        The date when the user completed work on the task.
        """
        return self._ensureSetNamed('_taskDateCompleted', '810F', constants.PSETID_TASK)

    @property
    def taskDeadOccurrence(self) -> bool:
        """
        Indicates whether a new recurring task remains to be generated. Set to
        False on a new Task object and True when the client generates the last
        recurring task.
        """
        return self._ensureSetNamed('_taskDeadOccurrence', '8109', constants.PSETID_TASK, overrideClass = bool, preserveNone = False)

    @property
    def taskDueDate(self) -> Optional[datetime.datetime]:
        """
        Specifies the date by which the user expects work on the task to be
        complete.
        """
        return self._ensureSetNamed('_taskDueDate', '8105', constants.PSETID_TASK)

    @property
    def taskEstimatedEffort(self) -> Optional[int]:
        """
        Indicates the number of minutes that the user expects to work on a task.
        """
        return self._ensureSetNamed('_taskEstimatedEffort', '8111', constants.PSETID_TASK)

    @property
    def taskFCreator(self) -> bool:
        """
        Indicates that the task object was originally created by the action of
        the current user or user agent instead of by the processing of a task
        request.
        """
        return self._ensureSetNamed('_taskFCreator', '811E', constants.PSETID_TASK, overrideClass = bool, preserveNone = False)

    @property
    def taskFFixOffline(self) -> bool:
        """
        Indicates whether the value of the taskOwner property is correct.
        """
        return self._ensureSetNamed('taskFFixOffline', '812C', constants.PSETID_TASK, overrideClass = bool, preserveNone = False)

    @property
    def taskFRecurring(self) -> bool:
        """
        Indicates whether the task includes a recurrence pattern.
        """
        return self._ensureSetNamed('_taskFRecurring', '8126', constants.PSETID_TASK, overrideClass = bool, preserveNone = False)

    @property
    def taskGlobalID(self) -> Optional[bytes]:
        """
        Specifies a unique GUID for this task, used to locate an existing task
        upon receipt of a task response or task update.
        """
        return self._ensureSetNamed('_taskGlobalID', '8519', constants.PSETID_COMMON)

    @property
    def taskHistory(self) -> Optional[TaskHistory]:
        """
        Indicates the type of change that was last made to the Task object.
        """
        return self._ensureSetNamed('_taskHistory', '811A', constants.PSETID_TASK, overrideClass = TaskHistory)

    @property
    def taskLastDelegate(self) -> Optional[str]:
        """
        Contains the name of the user who most recently assigned the task, or
        the user to whom it was most recently assigned.
        """
        return self._ensureSetNamed('_taskLastDelegate', '8125', constants.PSETID_TASK)

    @property
    def taskLastUpdate(self) -> Optional[datetime.datetime]:
        """
        The date and time of the most recent change made to the task object.
        """
        return self._ensureSetNamed('_taskLastUpdate', '8115', constants.PSETID_TASK)

    @property
    def taskLastUser(self) -> Optional[str]:
        """
        Contains the name of the most recent user to have been the owner of the
        task.
        """
        return self._ensureSetNamed('_taskLastUser', '8122', constants.PSETID_TASK)

    @property
    def taskMode(self) -> Optional[TaskMode]:
        """
        Used in a task communication. Should be 0 (UNASSIGNED) on task objects.
        """
        return self._ensureSetNamed('_taskMode', '8518', constants.PSETID_COMMON, overrideClass = TaskMode)

    @property
    def taskMultipleRecipients(self) -> Optional[Set[TaskMultipleRecipients]]:
        """
        Returns a set of flags that specify optimization hints about the
        recipients of a Task object.
        """
        return self._ensureSetNamed('_taskMultipleRecipients', '8120', constants.PSETID_TASK, overrideClass = TaskMultipleRecipients.fromBits)

    @property
    def taskNoCompute(self) -> Optional[bool]:
        """
        This value is not used and has no impact on a Task, but is provided for
        completeness.
        """
        return self._ensureSetNamed('_taskNoCompute', '8124', constants.PSETID_TASK)

    @property
    def taskOrdinal(self) -> Optional[int]:
        """
        Specifies a number that aids custom sorting of Task objects.
        """
        return self._ensureSetNamed('_taskOrdinal', '8123', constants.PSETID_TASK, overrideClass = unsignedToSignedInt)

    @property
    def taskOwner(self) -> Optional[str]:
        """
        Contains the name of the owner of the task.
        """
        return self._ensureSetNamed('_taskOwner', '811F', constants.PSETID_TASK)

    @property
    def taskOwnership(self) -> Optional[TaskOwnership]:
        """
        Contains the name of the owner of the task.
        """
        return self._ensureSetNamed('_taskOwnership', '8129', constants.PSETID_TASK, overrideClass = TaskOwnership)

    @property
    def taskRecurrence(self) -> Optional[RecurrencePattern]:
        """
        Contains a RecurrencePattern structure that provides information about
        recurring tasks.
        """
        return self._ensureSetNamed('_taskRecurrence', '8116', constants.PSETID_TASK, overrideClass = RecurrencePattern)

    @property
    def taskResetReminder(self) -> bool:
        """
        Indicates whether future recurring tasks need reminders.
        """
        return self._ensureSetNamed('_taskResetReminder', '8107', constants.PSETID_TASK, overrideClass = bool, preserveNone = False)

    @property
    def taskRole(self) -> Optional[str]:
        """
        This value is not used and has no impact on a Task, but is provided for
        completeness.
        """
        return self._ensureSetNamed('_taskRole', '8127', constants.PSETID_TASK)

    @property
    def taskStartDate(self) -> Optional[datetime.datetime]:
        """
        Specifies the date on which the user expects work on the task to begin.
        """
        return self._ensureSetNamed('_taskStartDate', '8104', constants.PSETID_TASK)

    @property
    def taskState(self) -> Optional[TaskState]:
        """
        Indicates the current assignment state of the Task object.
        """
        return self._ensureSetNamed('_taskState', '8113', constants.PSETID_TASK, overrideClass = TaskState)

    @property
    def taskStatus(self) -> Optional[TaskStatus]:
        """
        The completion status of a task.
        """
        return self._ensureSetNamed('_taskStatus', '8101', constants.PSETID_TASK, overrideClass = TaskStatus)

    @property
    def taskStatusOnComplete(self) -> bool:
        """
        Indicates whether the task assignee has been requested to send an email
        message upon completion of the assigned task.
        """
        return self._ensureSetNamed('_taskStatusOnComplete', '8119', constants.PSETID_TASK, overrideClass = bool, preserveNone = False)

    @property
    def taskUpdates(self) -> bool:
        """
        Indicates whether the task assignee has been requested to send a task
        update when the assigned Task object changes.
        """
        return self._ensureSetNamed('_taskUpdates', '811B', constants.PSETID_TASK, overrideClass = bool, preserveNone = False)

    @property
    def taskVersion(self) -> Optional[int]:
        """
        Indicates which copy is the latest update of a Task object.
        """
        return self._ensureSetNamed('_taskVersion', '8112', constants.PSETID_TASK)

    @property
    def teamTask(self) -> Optional[bool]:
        """
        This value is not used and has no impact on a Task, but is provided for
        completeness.
        """
        return self._ensureSetNamed('_teamTask', '8103', constants.PSETID_TASK)
