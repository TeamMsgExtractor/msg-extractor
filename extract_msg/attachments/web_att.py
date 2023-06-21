__all__ = [
    'WebAttachment',
]


from .. import constants
from .attachment_base import AttachmentBase
from ..enums import AttachmentPermissionType, AttachmentType


from typing import Optional


class WebAttachment(AttachmentBase):
    """
    An attachment that exists on the internet and not attached to the MSGFile
    directly.
    """

    def getFilename(self) -> str:
        raise NotImplementedError('Cannot get the filename of a web attachment.')

    def save(self, **kwargs) -> None:
        """
        Raises a NotImplementedError unless :param skipNotImplemented: is set to
        True. If it is, returns None to signify the attachment was skipped. This
        allows for the easy implementation of the option to skip this type of
        attachment.
        """
        if not kwargs.get('skipNotImplemented', False):
            raise NotImplementedError('Web attachments cannot be saved.')


    @property
    def data(self) -> None:
        """
        The bytes making up the attachment data.
        """
        raise NotImplementedError('Cannot get the data of a web attachment')

    @property
    def originalPermissionType(self) -> Optional[AttachmentPermissionType]:
        """
        The permission type data associated with a web reference attachment.
        """
        return self._ensureSetNamed('_oPermissionType', 'AttachmentOriginalPermissionType', constants.ps.PSETID_ATTACHMENT, overrideClass = AttachmentPermissionType, preserveNone = True)

    @property
    def permissionType(self) -> Optional[AttachmentPermissionType]:
        """
        The permission type data associated with a web reference attachment.
        """
        return self._ensureSetNamed('_permissionType', 'AttachmentPermissionType', constants.ps.PSETID_ATTACHMENT, overrideClass = AttachmentPermissionType, preserveNone = True)

    @property
    def providerName(self) -> Optional[str]:
        """
        The type of web service manipulating the attachment.
        """
        return self._ensureSetNamed('_permissionType', 'AttachmentProviderType', constants.ps.PSETID_ATTACHMENT)

    @property
    def type(self) -> AttachmentType:
        """
        Returns the (internally used) type of the data.
        """
        return AttachmentType.WEB

    @property
    def url(self) -> Optional[str]:
        """
        The url for the web attachment. If this is not set, that is probably an
        error.
        """
        return self.longPathname