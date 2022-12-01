import abc


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

    @property
    @abc.abstractmethod
    def data(self):
        """
        Gets the data for the attachment.
        """
