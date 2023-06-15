from __future__ import annotations


__all__ = [
    'openMsg',
    'openMsgBulk',
]


import glob
import logging

from typing import List, Tuple, TYPE_CHECKING, Union

from . import constants
from .exceptions import (
        InvalidFileFormatError, UnrecognizedMSGTypeError,
        UnsupportedMSGTypeError
    )


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

if TYPE_CHECKING:
    from .msg_classes import MSGFile


def _knownMsgClass(classType : str) -> bool:
    """
    Checks if the specified class type is recognized by the module. Usually used
    for checking if a type is simply unsupported rather than unknown.
    """
    classType = classType.lower()
    if classType == 'ipm':
        return True

    for item in constants.KNOWN_CLASS_TYPES:
        if classType.startswith(item):
            return True

    return False


def openMsg(path, **kwargs) -> MSGFile:
    """
    Function to automatically open an MSG file and detect what type it is.
    :param path: Path to the msg file in the system or is the raw msg file.
    :param prefix: Used for extracting embeded msg files inside the main one.
        Do not set manually unless you know what you are doing.
    :param parentMsg: Used for syncronizing named properties instances. Do not
        set this unless you know what you are doing.
    :param attachmentClass: Optional, the class the Message object will use for
        attachments. You probably should not change this value unless you know
        what you are doing.
    :param signedAttachmentClass: Optional, the class the object will use for
        signed attachments.
    :param filename: Optional, the filename to be used by default when saving.
    :param delayAttachments: Optional, delays the initialization of attachments
        until the user attempts to retrieve them. Allows MSG files with bad
        attachments to be initialized so the other data can be retrieved.
    :param overrideEncoding: Optional, overrides the specified encoding of the
        MSG file.
    :param errorBehavior: Optional, the behaviour to use in the event
        of an error when parsing the attachments.
    :param recipientSeparator: Optional, Separator string to use between
        recipients.
    :param ignoreRtfDeErrors: Optional, specifies that any errors that occur
        from the usage of RTFDE should be ignored (default: False).
    If :param strict: is set to `True`, this function will raise an exception
    when it cannot identify what MSGFile derivitive to use. Otherwise, it will
    log the error and return a basic MSGFile instance.
    :raises UnsupportedMSGTypeError: if the type is recognized but not suppoted.
    :raises UnrecognizedMSGTypeError: if the type is not recognized.
    """
    from .msg_classes import (
            AppointmentMeeting, Contact, MeetingCancellation, MeetingException,
            MeetingForwardNotification, MeetingRequest, MeetingResponse,
            Message, MSGFile, MessageSigned, Post, Task, TaskRequest
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

    # Restore the option in the kwargs so we don't have to worry about it.
    kwargs['delayAttachments'] = delayAttachments

    # After rechecking the docs, all comparisons should be case-insensitive, not
    # case-sensitive. My reading ability is great.
    #
    # Also after consideration, I realized we need to be very careful here, as
    # other file types (like doc, ppt, etc.) might open but not return a class
    # type. If the stream is not found, classType returns None, which has no
    # lower function. So let's make sure we got a good return first.
    if not msg.classType:
        if kwargs.get('strict', True):
            raise InvalidFileFormatError('File was confirmed to be an olefile, but was not an MSG file.')
        else:
            # If strict mode is off, we'll just return an MSGFile anyways.
            logger.critical('Received file that was an olefile but was not an MSG file. Returning MSGFile anyways because strict mode is off.')
            return msg
    classType = msg.classType.lower()
    # Put the message class first as it is most common.
    if classType.startswith('ipm.note') or classType.startswith('report'):
        msg.close()
        if classType.endswith('smime.multipartsigned') or classType.endswith('smime'):
            return MessageSigned(path, **kwargs)
        else:
            return Message(path, **kwargs)
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
        ct = msg.classType
        msg.close()
        if _knownMsgClass(classType):
            raise UnsupportedMSGTypeError(f'MSG type "{ct}" currently is not supported by the module. If you would like support, please make a feature request.')
        raise UnrecognizedMSGTypeError(f'Could not recognize msg class type "{ct}".')
    else:
        logger.error(f'Could not recognize msg class type "{msg.classType}". This most likely means it hasn\'t been implemented yet, and you should ask the developers to add support for it.')
        if not delayAttachments:
            msg.attachments
        return msg


def openMsgBulk(path, **kwargs) -> Union[List[MSGFile], Tuple[Exception, Union[str, bytes]]]:
    """
    Takes the same arguments as openMsg, but opens a collection of msg files
    based on a wild card. Returns a list if successful, otherwise returns a
    tuple.

    :param ignoreFailures: If this is True, will return a list of all successful
        files, ignoring any failures. Otherwise, will close all that
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