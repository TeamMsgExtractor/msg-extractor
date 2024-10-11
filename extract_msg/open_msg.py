from __future__ import annotations


__all__ = [
    'openMsg',
    'openMsgBulk',
]


import glob
import logging

from typing import List, Optional, Tuple, TYPE_CHECKING, Union

from . import constants
from .exceptions import (
        InvalidFileFormatError, UnrecognizedMSGTypeError,
        UnsupportedMSGTypeError
    )


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

if TYPE_CHECKING:
    from .msg_classes import MSGFile


def _getMsgClassInfo(classType: str) -> Tuple[bool, Optional[str]]:
    """
    Checks if the specified class type is recognized by the module.

    Usually used for checking if a type is simply unsupported rather than
    unknown.

    Returns a tuple of two items. The first is whether it is known. If it is
    known and support is refused, the second item will be a string of the
    relevent issue number. Otherwise, it will be None.
    """
    classType = classType.lower()
    if classType == 'ipm':
        return (True, None)

    for item in constants.KNOWN_CLASS_TYPES:
        if classType.startswith(item):
            # Check if the found class type has had support refused.
            for tup in constants.REFUSED_CLASS_TYPES:
                if tup[0] == item:
                    return (True, tup[1])
            else:
                return (True, None)

    return (False, None)


def openMsg(path, **kwargs) -> MSGFile:
    """
    Function to automatically open an MSG file and detect what type it is.

    Accepts all of the same arguments as the ``__init__`` method for the class
    it creates. Extra options will be ignored if the class doesn't know what to
    do with them, but child instances may end up using them if they understand
    them. See ``MSGFile.__init__`` for a list of all globally recognized
    options.

    :param strict: If set to ``True``, this function will raise an exception
        when it cannot identify what ``MSGFile`` derivitive to use. Otherwise,
        it will log the error and return a basic ``MSGFile`` instance. Default
        is ``True``.

    :raises UnsupportedMSGTypeError: The type is recognized but not suppoted.
    :raises UnrecognizedMSGTypeError: The type is not recognized.
    """
    from .msg_classes import (
            AppointmentMeeting, Contact, Journal, MeetingCancellation,
            MeetingException, MeetingForwardNotification, MeetingRequest,
            MeetingResponse, Message, MSGFile, MessageSigned, Post, StickyNote,
            Task, TaskRequest
        )

    # When the initial MSG file is opened, it should *always* delay attachments
    # so it can get the main class type. We only need to load them after that
    # if we are directly returning the MSGFile instance *and* delayAttachments
    # is False.
    #
    # So first let's store the original value.
    delayAttachments = kwargs.get('delayAttachments', False)
    kwargs['delayAttachments'] = True

    msg = MSGFile(path, **kwargs)

    # Protect against corrupt MSG files.
    try:
        ct = msg.classType
    except:
        # All exceptions here mean we NEED to close the handle.
        msg.close()
        raise

    # Restore the option in the kwargs so we don't have to worry about it.
    kwargs['delayAttachments'] = delayAttachments

    # After rechecking the docs, all comparisons should be case-insensitive, not
    # case-sensitive. My reading ability is great.
    #
    # Also after consideration, I realized we need to be very careful here, as
    # other file types (like doc, ppt, etc.) might open but not return a class
    # type. If the stream is not found, classType returns None, which has no
    # lower function. So let's make sure we got a good return first.
    if not ct:
        if kwargs.get('strict', True):
            raise InvalidFileFormatError('File was confirmed to be an olefile, but was not an MSG file.')
        else:
            # If strict mode is off, we'll just return an MSGFile anyways.
            logger.critical('Received file that was an olefile but was not an MSG file. Returning MSGFile anyways because strict mode is off.')
            return msg
    classType = ct.lower()
    # Put the message class first as it is most common.
    if classType.startswith('ipm.note') or classType.startswith('report') or classType.startswith('ipm.skypeteams.message'):
        msg.close()
        if classType.endswith('smime.multipartsigned') or classType.endswith('smime'):
            return MessageSigned(path, **kwargs)
        else:
            return Message(path, **kwargs)
    elif classType.startswith('ipm.activity'):
        msg.close()
        return Journal(path, **kwargs)
    elif classType.startswith('ipm.appointment'):
        msg.close()
        return AppointmentMeeting(path, **kwargs)
    elif classType.startswith('ipm.contact') or classType.startswith('ipm.distlist'):
        msg.close()
        return Contact(path, **kwargs)
    elif classType.startswith('ipm.post'):
        msg.close()
        return Post(path, **kwargs)
    elif classType.startswith('ipm.schedule.meeting.request'):
        msg.close()
        return MeetingRequest(path, **kwargs)
    elif classType.startswith('ipm.schedule.meeting.canceled'):
        msg.close()
        return MeetingCancellation(path, **kwargs)
    elif classType.startswith('ipm.schedule.meeting.notification.forward'):
        msg.close()
        return MeetingForwardNotification(path, **kwargs)
    elif classType.startswith('ipm.schedule.meeting.resp'):
        msg.close()
        return MeetingResponse(path, **kwargs)
    elif classType.startswith('ipm.stickynote'):
        msg.close()
        return StickyNote(path, **kwargs)
    elif classType.startswith('ipm.taskrequest'):
        msg.close()
        return TaskRequest(path, **kwargs)
    elif classType.startswith('ipm.task'):
        msg.close()
        return Task(path, **kwargs)
    elif classType.startswith('ipm.ole.class.{00061055-0000-0000-c000-000000000046}'):
        # Exception objects have a weird class type.
        msg.close()
        return MeetingException(path, **kwargs)
    elif classType == 'ipm':
        # Unspecified format. It should be equal to this and not just start with
        # it.
        if not delayAttachments:
            msg.attachments
        return msg
    elif kwargs.get('strict', True):
        # Because we are closing it, we need to store it in a variable first.
        msg.close()
        # Now we need to figure out exactly what we are going to be reporting to
        # the user.
        if (info := _getMsgClassInfo(classType))[0]:
            if info[1]:
                raise UnsupportedMSGTypeError(f'Support for MSG type "{ct}" has been refused. See {constants.REPOSITORY_URL}/issues/{info[1]} for more information.')
            raise UnsupportedMSGTypeError(f'MSG type "{ct}" currently is not supported by the module. If you would like support, please make a feature request.')
        raise UnrecognizedMSGTypeError(f'Could not recognize MSG class type "{ct}". As such, there is a high chance that support may be impossible, but you should contact the developers to find out more.')
    else:
        if (info := _getMsgClassInfo(classType))[0]:
            if info[1]:
                logger.error(f'Support for MSG type "{msg.classType}" has been refused. See {constants.REPOSITORY_URL}/issues/{info[1]} for more information.')
            else:
                logger.error(f'MSG type "{msg.classType}" currently is not supported by the module. If you would like support, please make a feature request.')
        logger.error(f'Could not recognize MSG class type "{msg.classType}". As such, there is a high chance that support may be impossible, but you should contact the developers to find out more.')
        if not delayAttachments:
            msg.attachments
        return msg


def openMsgBulk(path, **kwargs) -> Union[List[MSGFile], Tuple[Exception, Union[str, bytes]]]:
    """
    Takes the same arguments as openMsg, but opens a collection of MSG files
    based on a wild card. Returns a list if successful, otherwise returns a
    tuple.

    :param ignoreFailures: If this is ``True``, will return a list of all
        successful files, ignoring any failures. Otherwise, will close all that
        successfully opened, and return a tuple of the exception and the path of
        the file that failed.
    """
    files = []
    for x in glob.glob(str(path)):
        try:
            files.append(openMsg(x, **kwargs))
        except Exception as e:
            if not kwargs.get('ignoreFailures', False):
                for msg in files:
                    msg.close()
                return (e, x)

    return files