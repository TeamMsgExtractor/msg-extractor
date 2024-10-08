from __future__ import annotations


"""
Submodule for attachment classes.
"""

__all__ = [
    # Modules.
    'custom_att_handler',

    # Classes.
    'Attachment',
    'AttachmentBase',
    'BrokenAttachment',
    'CustomAttachmentHandler',
    'EmbeddedMsgAttachment',
    'SignedAttachment',
    'UnsupportedAttachment',
    'WebAttachment',

    # Functions.
    'initStandardAttachment',
    'registerHandler',
]


from . import custom_att_handler
from .attachment import Attachment
from .attachment_base import AttachmentBase
from .broken_att import BrokenAttachment
from .custom_att import CustomAttachment
from .custom_att_handler import CustomAttachmentHandler, registerHandler
from .emb_msg_att import EmbeddedMsgAttachment
from .signed_att import SignedAttachment
from .unsupported_att import UnsupportedAttachment
from .web_att import WebAttachment

import logging as _logging

from typing import TYPE_CHECKING


if TYPE_CHECKING:
    from ..msg_classes import MSGFile

_logger = _logging.getLogger(__name__)
_logger.addHandler(_logging.NullHandler())


def initStandardAttachment(msg: MSGFile, dir_: str) -> AttachmentBase:
    """
    Returns an instance of AttachmentBase for the attachment in the MSG file at
    the specified internal directory.

    :param errorBehavior: Used to tell the function what to do on errors.
    """
    from ..properties import PropertiesStore
    from ..enums import ErrorBehavior, PropertiesType
    from ..exceptions import (
            FeatureNotImplemented, StandardViolationError,
            UnrecognizedMSGTypeError
        )

    # First, create the properties store to check things like attachment type.
    propertiesStream = msg.getStream([dir_, '__properties_version1.0'])
    propStore = PropertiesStore(propertiesStream, PropertiesType.ATTACHMENT)

    try:
        # Now that we have the properties store, attempt to check what type it
        # is.
        if '37050003' not in propStore:
                from ..properties.prop import createProp

                _logger.warning(f'Attachment method property not found on attachment {dir_}. Code will attempt to guess the type.')
                _logger.log(5, propStore)

                # Because this condition is actually kind of a violation of the
                # standard, we are just going to do this in a dumb way.
                # Basically we are going to try to set the attach method
                # *manually* just so I don't have to go and modify the
                # following code.
                if msg.exists([dir_, '__substg1.0_37010102']):
                    # Set it as data and call it a day.
                    propData = b'\x03\x00\x057\x07\x00\x00\x00\x01\x00\x00\x00\x00\x00\x00\x00'
                elif msg.exists([dir_, '__substg1.0_3701000D']):
                    # If it is a folder and we have properties, call it an MSG
                    # file.
                    if msg.exists([dir_, '__substg1.0_3701000D/__properties_version1.0']):
                        propData = b'\x03\x00\x057\x07\x00\x00\x00\x05\x00\x00\x00\x00\x00\x00\x00'
                    else:
                        # Call if custom attachment data.
                        propData = b'\x03\x00\x057\x07\x00\x00\x00\x06\x00\x00\x00\x00\x00\x00\x00'
                else:
                    # Can't autodetect it, so throw an error.
                    raise StandardViolationError(f'Attachment method missing on attachment {dir_}, and it could not be determined automatically.')

                propStore.addProperty(createProp(propData), True)

        attMethod = propStore.getValue('37050003', 0) & 7

        if msg.exists([dir_, '__substg1.0_37010102']):
                return Attachment(msg, dir_, propStore)
        if msg.exists([dir_, '__substg1.0_3701000D']):
            if attMethod != 5:
                return CustomAttachment(msg, dir_, propStore)
            else:
                return EmbeddedMsgAttachment(msg, dir_, propStore)

        if attMethod == 7:
            return WebAttachment(msg, dir_, propStore)

        # Error handling.
        if attMethod == 1:
            raise StandardViolationError('Attachments of type data MUST have a data stream.')
        if attMethod == 0:
            raise NotImplementedError('extract-msg does not support attachments of type afNone.')
        if attMethod == 2:
            raise NotImplementedError('extract-msg does not support attachments of type afByReference. Contact the developer for support.')
        if attMethod == 4:
            raise NotImplementedError('extract-msg does not support attachments of type afByReferenceOnly. Contact the developer for support.')

        raise NotImplementedError(f'Could not determine attachment type ({attMethod})!')

    except (FeatureNotImplemented,
            NotImplementedError,
            UnrecognizedMSGTypeError):
        if ErrorBehavior.ATTACH_NOT_IMPLEMENTED in msg.errorBehavior:
            _logger.exception(f'Error processing attachment at {dir_}')
            return UnsupportedAttachment(msg, dir_, propStore)
        else:
            raise
    except StandardViolationError:
        if ErrorBehavior.STANDARDS_VIOLATION in msg.errorBehavior:
            _logger.exception(f'Unresolvable standards violation in  {dir_}')
            return BrokenAttachment(msg, dir_, propStore)
        else:
            raise
    except Exception:
        if ErrorBehavior.ATTACH_BROKEN in msg.errorBehavior:
            _logger.exception(f'Error processing attachment at {dir_}')
            return BrokenAttachment(msg, dir_, propStore)
        else:
            raise
