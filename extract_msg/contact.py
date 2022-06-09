from .msg import MSGFile


class Contact(MSGFile):
    """
    Class used for parsing contacts.
    """

    def __init__(self, path, **kwargs):
        """
        :param path: path to the msg file in the system or is the raw msg file.
        :param prefix: used for extracting embeded msg files
            inside the main one. Do not set manually unless
            you know what you are doing.
        :param attachmentClass: optional, the class the MSGFile object
            will use for attachments. You probably should
            not change this value unless you know what you
            are doing.
        :param delayAttachments: optional, delays the initialization of
            attachments until the user attempts to retrieve them. Allows MSG
            files with bad attachments to be initialized so the other data can
            be retrieved.
        :param filename: optional, the filename to be used by default when
            saving.
        :param overrideEncoding: optional, an encoding to use instead of the one
            specified by the msg file. Do not report encoding errors caused by
            this.
        """
        super().__init__(path, **kwargs)
        self.named
        self.namedProperties

    @property
    def birthday(self):
        """
        The birthday of the contact.
        """
        return self._ensureSetProperty('_birthday', '3A420040')

    @property
    def businessFax(self):
        """
        Contains the number of the contact's business fax.
        """
        return self._ensureSet('_businessFax', '__substg1.0_3A24')

    @property
    def businessPhone(self):
        """
        Contains the number of the contact's business phone.
        """
        return self._ensureSet('_businessPhone', '__substg1.0_3A08')

    @property
    def businessPhone2(self):
        """
        Contains the second number or numbers of the contact's
        business.
        """
        return self._ensureSetTyped('_businessPhone2', '3A1B')

    @property
    def businessUrl(self):
        """
        Contains the url of the homepage of the contact's business.
        """
        return self._ensureSet('_businessPhone', '__substg1.0_3A51')

    @property
    def callbackPhone(self):
        """
        Contains the number of the contact's car phone.
        """
        return self._ensureSet('_carPhone', '__substg1.0_3A1E')

    @property
    def callbackPhone(self):
        """
        Contains the contact's callback phone number.
        """
        return self._ensureSet('_callbackPhone', '__substg1.0_3A1E')

    @property
    def carPhone(self):
        """
        Contains the number of the contact's car phone.
        """
        return self._ensureSet('_carPhone', '__substg1.0_3A1E')

    @property
    def companyMainPhone(self):
        """
        Contains the number of the main phone of the contact's company.
        """
        return self._ensureSet('_businessPhone', '__substg1.0_3A57')

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
        The name of the department the contact works in.
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
    def instantMessagingAddress(self):
        """
        The instant messaging address of the contact.
        """
        return self._ensureSetNamed('_instantMessagingAddress', '8062')

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
    def spouseName(self):
        """
        The name of the contact's spouse.
        """
        return self._ensureSet('_spouseName', '__substg1.0_3A')

    @property
    def state(self):
        """
        The state or province that the contact lives in.
        """
        return self._ensureSet('_state', '__substg1.0_3A28')

    @property
    def workAddress(self):
        """
        The work address of the contact.
        """
        return self._ensureSetNamed('_workAddress', '801B')
