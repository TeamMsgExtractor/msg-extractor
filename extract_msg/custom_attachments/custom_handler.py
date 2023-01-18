import abc

from typing import List, Optional, Tuple


class CustomAttachmentHandler(abc.ABC):
    """
    A class designed to help with custom attachments that may require parsing in
    special ways that are completely different from one another.
    """
    def __init__(self, attachment : 'Attachment'):
        super().__init__()
        self.__att = attachment

    @classmethod
    @abc.abstractmethod
    def isCorrectHandler(cls, attachment : 'Attachment') -> bool:
        """
        Checks if this is the correct handler for the attachment.
        """

    @abc.abstractmethod
    def injectHTML(self, html : bytes, renderedList : Optional[List[str]] = None) -> Tuple[bytes, Optional[List[str]]]:
        """
        Adds the relevent tag, if any, to the HTML for making prepared HTML.

        If this function should do nothing, returns the two arguments without
        modification.

        :param html: The HTML body to inject into (if at all).
        :param renderedList: The list to use (if needed) of "rendered
            characters" which will be returned and matches the HTML returned.
        """

    @property
    def attachment(self):
        """
        The attachment this handler is associated with.
        """
        return self.__att

    @property
    @abc.abstractmethod
    def data(self) -> bytes:
        """
        Gets the data for the attachment.
        """

    @property
    @abc.abstractmethod
    def name(self) -> str:
        """
        Returns the name to be used when saving the attachment.
        """
