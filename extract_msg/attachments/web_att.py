__all__ = [
    'WebAttachment',
]


import functools

from typing import Optional

from .. import constants
from .attachment_base import AttachmentBase
from ..enums import AttachmentPermissionType, AttachmentType, SaveType


class WebAttachment(AttachmentBase):
    """
    An attachment that exists on the internet and not attached to the MSG file
    directly.
    """

    def getFilename(self) -> str:
        raise NotImplementedError('Cannot get the filename of a web attachment.')

    def save(self, **kwargs) -> constants.SAVE_TYPE:
        """
        Raises a NotImplementedError unless :param skipNotImplemented: is set to
        True.

        If it is, returns a value that indicates no data was saved.
        """
        if kwargs.get('skipNotImplemented', False):
            return (SaveType.NONE, None)

        raise NotImplementedError('Web attachments cannot be saved.')

    @property
    def data(self) -> None:
        """
        The bytes making up the attachment data.
        """
        raise NotImplementedError('Cannot get the data of a web attachment.')

    @functools.cached_property
    def originalPermissionType(self) -> Optional[AttachmentPermissionType]:
        """
        The permission type data associated with a web reference attachment.
        """
        return self.getNamedAs('AttachmentOriginalPermissionType', constants.ps.PSETID_ATTACHMENT, AttachmentPermissionType)

    @functools.cached_property
    def permissionType(self) -> Optional[AttachmentPermissionType]:
        """
        The permission type data associated with a web reference attachment.
        """
        return self.getNamedAs('AttachmentPermissionType', constants.ps.PSETID_ATTACHMENT, AttachmentPermissionType)

    @functools.cached_property
    def providerName(self) -> Optional[str]:
        """
        The type of web service manipulating the attachment.
        """
        return self.getNamedProp('AttachmentProviderType', constants.ps.PSETID_ATTACHMENT)

    @property
    def type(self) -> AttachmentType:
        return AttachmentType.WEB

    @property
    def url(self) -> Optional[str]:
        """
        The url for the web attachment. If this is not set, that is probably an
        error.
        """
        return self.longPathname