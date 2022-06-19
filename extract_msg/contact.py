import datetime

from .enums import Gender
from .msg import MSGFile


class Contact(MSGFile):
    """
    Class used for parsing contacts.
    """

    def __init__(self, path, **kwargs):
        """
        :param path: path to the msg file in the system or is the raw msg file.
        :param prefix: used for extracting embeded msg files inside the main
            one. Do not set manually unless you know what you are doing.
        :param attachmentClass: optional, the class the MSGFile object will use
            for attachments. You probably should not change this value unless
            you know what you are doing.
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

    def save(self, **kwargs):
        """
        Save function.
        """
        # TODO.
        pass

    @property
    def birthday(self) -> datetime.datetime:
        """
        The birthday of the contact.
        """
        return self._ensureSetProperty('_birthday', '3A420040')

    @property
    def businessFax(self) -> str:
        """
        Contains the number of the contact's business fax.
        """
        return self._ensureSet('_businessFax', '__substg1.0_3A24')

    @property
    def businessPhone(self) -> str:
        """
        Contains the number of the contact's business phone.
        """
        return self._ensureSet('_businessPhone', '__substg1.0_3A08')

    @property
    def businessPhone2(self) -> str:
        """
        Contains the second number or numbers of the contact's business.
        """
        return self._ensureSetTyped('_businessPhone2', '3A1B')

    @property
    def businessUrl(self) -> str:
        """
        Contains the url of the homepage of the contact's business.
        """
        return self._ensureSet('_businessPhone', '__substg1.0_3A51')

    @property
    def callbackPhone(self) -> str:
        """
        Contains the number of the contact's car phone.
        """
        return self._ensureSet('_carPhone', '__substg1.0_3A1E')

    @property
    def callbackPhone(self) -> str:
        """
        Contains the contact's callback phone number.
        """
        return self._ensureSet('_callbackPhone', '__substg1.0_3A1E')

    @property
    def carPhone(self) -> str:
        """
        Contains the number of the contact's car phone.
        """
        return self._ensureSet('_carPhone', '__substg1.0_3A1E')

    @property
    def companyMainPhone(self) -> str:
        """
        Contains the number of the main phone of the contact's company.
        """
        return self._ensureSet('_businessPhone', '__substg1.0_3A57')

    @property
    def companyName(self) -> str:
        """
        The name of the company the contact works at.
        """
        return self._ensureSet('_companyName', '__substg1.0_3A16')

    @property
    def country(self) -> str:
        """
        The country the contact lives in.
        """
        return self._ensureSet('_country', '__substg1.0_3A26')

    @property
    def departmentName(self) -> str:
        """
        The name of the department the contact works in.
        """
        return self._ensureSet('_departmentName', '__substg1.0_3A18')

    @property
    def displayName(self) -> str:
        """
        The full name of the contact.
        """
        return self._ensureSet('_displayName', '__substg1.0_3001')

    @property
    def displayNamePrefix(self) -> str:
        """
        The contact's honorific title.
        """
        return self._ensureSet('_displayNamePrefix', '__substg1.0_3A45')

    def email1(self) -> dict:
        """
        Returns a dict of the data for email 1.
        """
        return {
            'address_type': self.email1AddressType,
            'display_name': self.email1DisplayName,
            'email_address': self.email1EmailAddress,
            'original_display_name': self.email1OriginalDisplayName,
            'original_entry_id': self.email1OriginalEntryId,
        }

    @property
    def email1AddressType(self) -> str:
        """
        The address type of the first email address.
        """
        return self._ensureSetNamed('_email1AddressType', '8082')

    @property
    def email1DisplayName(self) -> str:
        """
        The user-readable display name of the first email address.
        """
        return self._ensureSetNamed('_email1DisplayName', '8080')

    @property
    def email1EmailAddress(self) -> str:
        """
        The first email address.
        """
        return self._ensureSetNamed('_email1EmailAddress', '8083')

    @property
    def email1OriginalDisplayName(self) -> str:
        """
        The first SMTP email address that corresponds to the first email address
        for the contact.
        """
        return self._ensureSetNamed('_email1OriginalDisplayName', '8084')

    @property
    def email1OriginalEntryId(self) -> EntryID:
        """
        The EntryID of the object correspinding to this electronic address.
        """
        return self._ensureSetNamed('_email1OriginalEntryId', '8085', overrideClass = EntryID.autoCreate)

    def email2(self) -> dict:
        """
        Returns a dict of the data for email 2.
        """
        return {
            'address_type': self.email2AddressType,
            'display_name': self.email2DisplayName,
            'email_address': self.email2EmailAddress,
            'original_display_name': self.email2OriginalDisplayName,
            'original_entry_id': self.email2OriginalEntryId,
        }

    @property
    def email2AddressType(self) -> str:
        """
        The address type of the second email address.
        """
        return self._ensureSetNamed('_email2AddressType', '8092')

    @property
    def email2DisplayName(self) -> str:
        """
        The user-readable display name of the second email address.
        """
        return self._ensureSetNamed('_email2DisplayName', '8090')

    @property
    def email2EmailAddress(self) -> str:
        """
        The second email address.
        """
        return self._ensureSetNamed('_email2EmailAddress', '8093')

    @property
    def email2OriginalDisplayName(self) -> str:
        """
        The second SMTP email address that corresponds to the second email address
        for the contact.
        """
        return self._ensureSetNamed('_email2OriginalDisplayName', '8094')

    @property
    def email2OriginalEntryId(self) -> EntryID:
        """
        The EntryID of the object correspinding to this electronic address.
        """
        return self._ensureSetNamed('_email2OriginalEntryId', '8095', overrideClass = EntryID.autoCreate)

    @property
    def email3(self) -> dict:
        """
        Returns a dict of the data for email 3.
        """
        return {
            'address_type': self.email3AddressType,
            'display_name': self.email3DisplayName,
            'email_address': self.email3EmailAddress,
            'original_display_name': self.email3OriginalDisplayName,
            'original_entry_id': self.email3OriginalEntryId,
        }

    @property
    def email3AddressType(self) -> str:
        """
        The address type of the third email address.
        """
        return self._ensureSetNamed('_email3AddressType', '80A2')

    @property
    def email3DisplayName(self) -> str:
        """
        The user-readable display name of the third email address.
        """
        return self._ensureSetNamed('_email3DisplayName', '80A0')

    @property
    def email3EmailAddress(self) -> str:
        """
        The third email address.
        """
        return self._ensureSetNamed('_email3EmailAddress', '80A3')

    @property
    def email3OriginalDisplayName(self) -> str:
        """
        The third SMTP email address that corresponds to the third email address
        for the contact.
        """
        return self._ensureSetNamed('_email3OriginalDisplayName', '80A4')

    @property
    def email3OriginalEntryId(self) -> EntryID:
        """
        The EntryID of the object correspinding to this electronic address.
        """
        return self._ensureSetNamed('_email3OriginalEntryId', '80A5', overrideClass = EntryID.autoCreate)

    @property
    def fileUnder(self) -> str:
        """
        The name under which to file a contact when displaying a list of
        contacts.
        """
        return self._ensureSetNamed('_fileUnder', '8005')

    @property
    def fileUnderID(self) -> int:
        """
        The format to use for fileUnder. See PidLidFileUnderId in [MS-OXOCNTC]
        for details.
        """
        return self._ensureSetNamed('_fileUnderID', '8006')

    @property
    def gender(self) -> Gender:
        """
        The gender of the contact.
        """

    @property
    def generation(self):
        """
        A generational abbreviation that follows the full
        name of the contact.
        """
        return self._ensureSet('_generation', '__substg1.0_3A05')

    @property
    def givenName(self):
        """
        The first name of the contact.
        """
        return self._ensureSet('_givenName', '__substg1.0_3A06')

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
    def nickname(self):
        """
        The nickname of the contanct.
        """
        return self._ensureSet('_nickname', '__substg1.0_3A4F')

    @property
    def phoneticGivenName(self):
        """
        The phonetic pronunciation of the given name of the contact.
        """
        return self._ensureSetNamed('_phoneticGivenName', '802C')

    @property
    def phoneticSurname(self):
        """
        The phonetic pronunciation of the given name of the contact.
        """
        return self._ensureSetNamed('_phoneticSurname', '802D')

    @property
    def spouseName(self):
        """
        The name of the contact's spouse.
        """
        return self._ensureSet('_spouseName', '__substg1.0_3A48')

    @property
    def state(self):
        """
        The state or province that the contact lives in.
        """
        return self._ensureSet('_state', '__substg1.0_3A28')

    @property
    def surname(self):
        """
        The surname of the contact.
        """
        return self._ensureSet('_surname', '__substg1.0_3A11')

    @property
    def workAddress(self):
        """
        The work address of the contact.
        """
        return self._ensureSetNamed('_workAddress', '801B')
