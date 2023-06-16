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

    def save(self, **_) -> None:
        raise NotImplementedError('Cannot save a web attachment.')

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
        return self._ensureSetNamed('_oPermissionType', 'AttachmentOriginalPermissionType', constants.PSETID_ATTACHMENT, overrideClass = AttachmentPermissionType, preserveNone = True)

    @property
    def permissionType(self) -> Optional[AttachmentPermissionType]:
        """
        The permission type data associated with a web reference attachment.
        """
        return self._ensureSetNamed('_permissionType', 'AttachmentPermissionType', constants.PSETID_ATTACHMENT, overrideClass = AttachmentPermissionType, preserveNone = True)

    @property
    def providerName(self) -> Optional[str]:
        """
        The type of web service manipulating the attachment.
        """
        return self._ensureSetNamed('_permissionType', 'AttachmentProviderType', constants.PSETID_ATTACHMENT)

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