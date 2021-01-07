from extract_msg import constants
from extract_msg.attachment import Attachment
from extract_msg.msg import MSGFile


class Contact(MSGFile):
    """
    Class used for parsing contacts.
    """

    def __init__(self, path, prefix = '', attachmentClass = Attachment, filename = None, overrideEncoding = None, attachmentErrorBehavior = constants.ATTACHMENT_ERROR_THROW):
        MSGFile.__init__(self, path, prefix, attachmentClass, filename, overrideEncoding, attachmentErrorBehavior)
        self.named

    @property
    def businessPhone(self):
        """
        Contains the number of the contact's business phone.
        """
        return self._ensureSet('_businessPhone', '__substg1.0_3A08')

    @property
    def companyName(self):
        """
        The name of the company the contact works at.
        """
        return self._ensureSet('_companyName', '__substg1.0_3A16')

    @property
    def country(self):
        """
        The country the contact lives in.
        """
        return self._ensureSet('_country', '__substg1.0_3A26')

    @property
    def departmentName(self):
        """
        The name of the dapartment the contact works in.
        """
        return self._ensureSet('_departmentName', '__substg1.0_3A18')

    @property
    def firstName(self):
        """
        The first name of the contact.
        """
        return self._ensureSet('_firstName', '__substg1.0_3A06')

    @property
    def generation(self):
        """
        A generational abbreviation that follows the full
        name of the contact.
        """
        return self._ensureSet('_generation', '__substg1.0_3A05')

    @property
    def honorificTitle(self):
        """
        The honorific title of the contact.
        """
        return self._ensureSet('_honorificTitle', '__substg1.0_3A45')

    @property
    def initials(self):
        """
        The initials of the contact.
        """
        return self._ensureSet('_initials', '__substg1.0_3A0A')

    @property
    def jobTitle(self):
        """
        The job title of the contact.
        """
        return self._ensureSet('_jobTitle', '__substg1.0_3A17')

    @property
    def lastModifiedBy(self):
        """
        The name of the last user to modify the contact file.
        """
        return self._ensureSet('_lastModifiedBy', '__substg1.0_3FFA')

    @property
    def lastName(self):
        """
        The last name of the contact.
        """
        return self._ensureSet('_lastName', '__substg1.0_3A11')

    @property
    def locality(self):
        """
        The locality (such as town or city) of the contact.
        """
        return self._ensureSet('_locality', '__substg1.0_3A27')

    @property
    def middleNames(self):
        """
        The middle name(s) of the contact.
        """
        return self._ensureSet('_middleNames', '__substg1.0_3A44')

    @property
    def mobilePhone(self):
        """
        The mobile phone number of the contact.
        """
        return self._ensureSet('_mobilePhone', '__substg1.0_3A1C')

    @property
    def state(self):
        """
        The state or province that the contact lives in.
        """
        return self._ensureSet('_state', '__substg1.0_3A28')
