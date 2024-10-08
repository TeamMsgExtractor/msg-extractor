__all__ = [
    'Task',
]


import datetime
import functools
import json
import logging

from typing import Optional

from .. import constants
from ..enums import (
        TaskAcceptance, TaskHistory, TaskMode, TaskMultipleRecipients,
        TaskOwnership, TaskState, TaskStatus
    )
from .message_base import MessageBase
from ..structures.recurrence_pattern import RecurrencePattern
from ..utils import unsignedToSignedInt


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Task(MessageBase):
    """
    Class used for parsing task files.
    """

    def getJson(self) -> str:
        status = {
            TaskStatus.NOT_STARTED: 'Not Started',
            TaskStatus.IN_PROGRESS: 'In Progress',
            TaskStatus.COMPLETE: 'Completed',
            TaskStatus.WAITING_ON_OTHER: 'Waiting on someone else',
            TaskStatus.DEFERRED: 'Deferred',
            None: None,
        }[self.taskStatus]

        return json.dumps({
            'subject': self.subject,
            'status': status,
            'percentComplete': f'{self.percentComplete*100:.0f}%',
            'dateCompleted': self.taskDateCompleted.__format__(self.dateFormat) if self.taskDateCompleted else None,
            'totalWork': f'{self.taskEstimatedEffort or 0} minutes',
            'actualWork': f'{self.taskActualEffort or 0} minutes',
            'owner': self.taskOwner,
            'importance': self.importanceString,
        })

    @property
    def headerFormatProperties(self) -> constants.HEADER_FORMAT_TYPE:
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
                'Date Completed': self.taskDateCompleted.__format__(self.dateFormat) if self.taskDateCompleted else None,
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

    @functools.cached_property
    def percentComplete(self) -> float:
        """
        Indicates whether a time-flagged Message object is complete.

        :returns: A percentage in decimal form. 1.0 indicates it is complete.
        """
        return self.getNamedProp('8102', constants.ps.PSETID_TASK, 0.0)

    @functools.cached_property
    def taskAcceptanceState(self) -> Optional[TaskAcceptance]:
        """
        Indicates the acceptance state of the task.
        """
        return self.getNamedAs('812A', constants.ps.PSETID_TASK, TaskAcceptance)

    @functools.cached_property
    def taskAccepted(self) -> bool:
        """
        Indicates whether a task assignee has replied to a tesk request for this
        task object.

        Does not indicate if it was accepted or rejected.
        """
        return bool(self.getNamedProp('8108', constants.ps.PSETID_TASK))

    @functools.cached_property
    def taskActualEffort(self) -> Optional[int]:
        """
        Indicates the number of minutes that the user actually spent working on
        a task.
        """
        return self.getNamedProp('8110', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskAssigner(self) -> Optional[str]:
        """
        Specifies the name of the user that last assigned the task.
        """
        return self.getNamedProp('8121', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskAssigners(self) -> Optional[bytes]:
        """
        A stack of entries, each representing a task assigner.

        The most recent task assigner (that is, the top) appears at the bottom.

        The documentation on this is weird, so I don't know how to parse it.
        """
        return self.getNamedProp('8117', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskComplete(self) -> bool:
        """
        Indicates if the task is complete.
        """
        return bool(self.getNamedProp('811C', constants.ps.PSETID_TASK))

    @functools.cached_property
    def taskCustomFlags(self) -> Optional[int]:
        """
        Custom flags set on the task.
        """
        return self.getNamedProp('8139', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskDateCompleted(self) -> Optional[datetime.datetime]:
        """
        The date when the user completed work on the task.
        """
        return self.getNamedProp('810F', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskDeadOccurrence(self) -> bool:
        """
        Indicates whether a new recurring task remains to be generated.

        Set to ``False`` on a new Task object and ``True`` when the client
        generates the last recurring task.
        """
        return bool(self.getNamedProp('8109', constants.ps.PSETID_TASK))

    @functools.cached_property
    def taskDueDate(self) -> Optional[datetime.datetime]:
        """
        Specifies the date by which the user expects work on the task to be
        complete.
        """
        return self.getNamedProp('8105', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskEstimatedEffort(self) -> Optional[int]:
        """
        Indicates the number of minutes that the user expects to work on a task.
        """
        return self.getNamedProp('8111', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskFCreator(self) -> bool:
        """
        Indicates that the task object was originally created by the action of
        the current user or user agent instead of by the processing of a task
        request.
        """
        return bool(self.getNamedProp('811E', constants.ps.PSETID_TASK))

    @functools.cached_property
    def taskFFixOffline(self) -> bool:
        """
        Indicates whether the value of the taskOwner property is correct.
        """
        return bool(self.getNamedProp('812C', constants.ps.PSETID_TASK))

    @functools.cached_property
    def taskFRecurring(self) -> bool:
        """
        Indicates whether the task includes a recurrence pattern.
        """
        return bool(self.getNamedProp('8126', constants.ps.PSETID_TASK))

    @functools.cached_property
    def taskGlobalID(self) -> Optional[bytes]:
        """
        Specifies a unique GUID for this task, used to locate an existing task
        upon receipt of a task response or task update.
        """
        return self.getNamedProp('8519', constants.ps.PSETID_COMMON)

    @functools.cached_property
    def taskHistory(self) -> Optional[TaskHistory]:
        """
        Indicates the type of change that was last made to the Task object.
        """
        return self.getNamedAs('811A', constants.ps.PSETID_TASK, TaskHistory)

    @functools.cached_property
    def taskLastDelegate(self) -> Optional[str]:
        """
        Contains the name of the user who most recently assigned the task, or
        the user to whom it was most recently assigned.
        """
        return self.getNamedProp('8125', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskLastUpdate(self) -> Optional[datetime.datetime]:
        """
        The date and time of the most recent change made to the task object.
        """
        return self.getNamedProp('8115', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskLastUser(self) -> Optional[str]:
        """
        Contains the name of the most recent user to have been the owner of the
        task.
        """
        return self.getNamedProp('8122', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskMode(self) -> Optional[TaskMode]:
        """
        Used in a task communication.

        Should be 0 (UNASSIGNED) on task objects.
        """
        return self.getNamedAs('8518', constants.ps.PSETID_COMMON, TaskMode)

    @functools.cached_property
    def taskMultipleRecipients(self) -> Optional[TaskMultipleRecipients]:
        """
        Returns a union of flags that specify optimization hints about the
        recipients of a Task object.
        """
        return self.getNamedAs('8120', constants.ps.PSETID_TASK, TaskMultipleRecipients)

    @functools.cached_property
    def taskNoCompute(self) -> Optional[bool]:
        """
        This value is not used and has no impact on a Task, but is provided for
        completeness.
        """
        return self.getNamedProp('8124', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskOrdinal(self) -> Optional[int]:
        """
        Specifies a number that aids custom sorting of Task objects.
        """
        return self.getNamedAs('8123', constants.ps.PSETID_TASK, unsignedToSignedInt)

    @functools.cached_property
    def taskOwner(self) -> Optional[str]:
        """
        Contains the name of the owner of the task.
        """
        return self.getNamedProp('811F', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskOwnership(self) -> Optional[TaskOwnership]:
        """
        Contains the name of the owner of the task.
        """
        return self.getNamedAs('8129', constants.ps.PSETID_TASK, TaskOwnership)

    @functools.cached_property
    def taskRecurrence(self) -> Optional[RecurrencePattern]:
        """
        Contains a RecurrencePattern structure that provides information about
        recurring tasks.
        """
        return self.getNamedAs('8116', constants.ps.PSETID_TASK, RecurrencePattern)

    @functools.cached_property
    def taskResetReminder(self) -> bool:
        """
        Indicates whether future recurring tasks need reminders.
        """
        return bool(self.getNamedProp('8107', constants.ps.PSETID_TASK))

    @functools.cached_property
    def taskRole(self) -> Optional[str]:
        """
        This value is not used and has no impact on a Task, but is provided for
        completeness.
        """
        return self.getNamedProp('8127', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskStartDate(self) -> Optional[datetime.datetime]:
        """
        Specifies the date on which the user expects work on the task to begin.
        """
        return self.getNamedProp('8104', constants.ps.PSETID_TASK)

    @functools.cached_property
    def taskState(self) -> Optional[TaskState]:
        """
        Indicates the current assignment state of the Task object.
        """
        return self.getNamedAs('8113', constants.ps.PSETID_TASK, TaskState)

    @functools.cached_property
    def taskStatus(self) -> Optional[TaskStatus]:
        """
        The completion status of a task.
        """
        return self.getNamedAs('8101', constants.ps.PSETID_TASK, TaskStatus)

    @functools.cached_property
    def taskStatusOnComplete(self) -> bool:
        """
        Indicates whether the task assignee has been requested to send an email
        message upon completion of the assigned task.
        """
        return bool(self.getNamedProp('8119', constants.ps.PSETID_TASK))

    @functools.cached_property
    def taskUpdates(self) -> bool:
        """
        Indicates whether the task assignee has been requested to send a task
        update when the assigned Task object changes.
        """
        return bool(self.getNamedProp('811B', constants.ps.PSETID_TASK))

    @functools.cached_property
    def taskVersion(self) -> Optional[int]:
        """
        Indicates which copy is the latest update of a Task object.
        """
        return self.getNamedProp('8112', constants.ps.PSETID_TASK)

    @functools.cached_property
    def teamTask(self) -> Optional[bool]:
        """
        This value is not used and has no impact on a Task, but is provided for
        completeness.
        """
        return self.getNamedProp('8103', constants.ps.PSETID_TASK)
