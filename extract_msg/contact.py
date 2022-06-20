import datetime

from typing import Tuple, Set

from .enums import ElectronicAddressProperties, Gender, PostalAddressID
from .msg import MSGFile
from .structures.entry_id import EntryID


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
    def addressBookProviderArrayType(self) -> Set[ElectronicAddressProperties]:
        """
        A set of which Electronic Address properties are set on the contact.

        Property is stored in the MSG file as a sinlge int. The result should be
        identical to addressBookProviderEmailList.
        """
        return self._ensureSetNamed('_addressBookProviderArrayType', '8029', ElectronicAddressProperties.fromBits)

    @property
    def addressBookProviderEmailList(self) -> Set[ElectronicAddressProperties]:
        """
        A set of which Electronic Address properties are set on the contact.
        """
        return self._ensureSetNamed('_addressBookProviderEmailList', '8028', overrideClass = lambda x : {ElectronicAddressProperties(y) for y in x})

    @property
    def birthday(self) -> datetime.datetime:
        """
        The birthday of the contact.
        """
        return self._ensureSetProperty('_birthday', '3A420040')

    @property
    def businessFax(self) -> dict:
        """
        Returns a dict of the data for the business fax. Returns None if no
        fields are set.

        Keys are "address_type", "email_address", "number",
        "original_display_name", and "original_entry_id".
        """
        try:
            return self._businessFax
        except AttributeError:
            data = {
                'address_type': self.businessFaxAddressType,
                'email_address': self.businessFaxEmailAddress,
                'number': self.businessFaxNumber,
                'original_display_name': self.businessFaxOriginalDisplayName,
                'original_entry_id': self.businessFaxOriginalEntryId,
            }
            self._businessFax = data if any(data[x] for x in data) else None
            return self._businessFax

    @property
    def businessFaxAddressType(self) -> str:
        """
        The type of address for the fax. MUST be set to "FAX" if present.
        """
        return self._ensureSetNamed('_businessFaxAddressType', '80C2')

    @property
    def businessFaxEmailAddress(self) -> str:
        """
        Contains a user-readable display name, followed by the "@" character,
        followed by a fax number.
        """
        return self._ensureSetNamed('_businessFaxAddressType', '80C3')

    @property
    def businessFaxNumber(self) -> str:
        """
        Contains the number of the contact's business fax.
        """
        return self._ensureSet('_businessFaxNumber', '__substg1.0_3A24')

    @property
    def businessFaxOriginalDisplayName(self) -> str:
        """
        The normalized subject for the contact.
        """
        return self._ensureSetNamed('_businessFaxAddressType', '80C4')

    @property
    def businessFaxOriginalEntryId(self) -> EntryID:
        """
        The one-off EntryID corresponding to this fax address.
        """
        return self._ensureSetNamed('_businessFaxAddressType', '80C5', overrideClass = EntryID.autoCreate)

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
        Contains the contact's callback phone number.
        """
        return self._ensureSet('_callbackPhone', '__substg1.0_3A02')

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

    @property
    def email1(self) -> dict:
        """
        Returns a dict of the data for email 1. Returns None if no fields are
        set.

        Keys are "address_type", "display_name", "email_address",
        "original_display_name", and "original_entry_id".
        """
        try:
            return self._email1
        except AttributeError:
            data = {
                'address_type': self.email1AddressType,
                'display_name': self.email1DisplayName,
                'email_address': self.email1EmailAddress,
                'original_display_name': self.email1OriginalDisplayName,
                'original_entry_id': self.email1OriginalEntryId,
            }
            self._email1 = data if any(data[x] for x in data) else None
            return self._email1

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

    @property
    def email2(self) -> dict:
        """
        Returns a dict of the data for email 2. Returns None if no fields are
        set.
        """
        try:
            return self._email2
        except AttributeError:
            data = {
                'address_type': self.email2AddressType,
                'display_name': self.email2DisplayName,
                'email_address': self.email2EmailAddress,
                'original_display_name': self.email2OriginalDisplayName,
                'original_entry_id': self.email2OriginalEntryId,
            }
            self._email2 = data if any(data[x] for x in data) else None
            return self._email2

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
        Returns a dict of the data for email 3. Returns None if no fields are
        set.
        """
        try:
            return self._email3
        except AttributeError:
            data = {
                'address_type': self.email3AddressType,
                'display_name': self.email3DisplayName,
                'email_address': self.email3EmailAddress,
                'original_display_name': self.email3OriginalDisplayName,
                'original_entry_id': self.email3OriginalEntryId,
            }
            self._email3 = data if any(data[x] for x in data) else None
            return self._email3

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
    def emails(self) -> tuple:
        """
        Returns a tuple of all the email dicts. Value for an email will be None
        if no fields were set.
        """
        try:
            return self._emails
        except AttributeError:
            self._emails = (self.email1, self.email2, self.email3)
            return self._emails

    @property
    def faxNumbers(self) -> dict:
        """
        Returns a dictionary of the fax numbers. Entry will be None if no fields
        were set.

        Keys are "business", "home", and "primary".
        """
        try:
            return self._faxNumbers
        except AttributeError:
            self._faxNumbers = {
                'business': self.businessFax,
                'home': self.homeFax,
                'primary': self.primaryFax,
            }

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
        return self._ensureSet('_gender', '__substg1.0_3A4D')

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
    def homeAddress(self) -> str:
        """
        The complete home address of the contact.
        """
        return self._ensureSetNamed('_homeAddress', '801A')

    @property
    def homeAddressCountry(self) -> str:
        """
        The country portion of the contact's home address.
        """
        return self._ensureSet('_homeAddressCountry', '__substg1.0_3A5A')

    @property
    def homeAddressCountryCode(self) -> str:
        """
        The country code portion of the contact's home address.
        """
        return self._ensureSetNamed('_homeAddressCountryCode', '80DA')

    @property
    def homeAddressLocality(self) -> str:
        """
        The locality or city portion of the contact's home address.
        """
        return self._ensureSet('_homeAddressLocality', '__substg1.0_3A59')

    @property
    def homeAddressPostalCode(self) -> str:
        """
        The postal code portion of the contact's home address.
        """
        return self._ensureSet('_homeAddressPostalCode', '__substg1.0_3A5B')

    @property
    def homeAddressPostOfficeBox(self) -> str:
        """
        The number or identifier of the contact's home post office box.
        """
        return self._ensureSet('_homeAddressPostOfficeBox', '__substg1.0_3A5E')

    @property
    def homeAddressStateOrProvince(self) -> str:
        """
        The state or province portion of the contact's home address.
        """
        return self._ensureSet('_homeAddressStateOrProvince', '__substg1.0_3A5C')

    @property
    def homeAddressStreet(self) -> str:
        """
        The street portion of the contact's home address.
        """
        return self._ensureSet('_homeAddressStreet', '__substg1.0_3A5D')

    @property
    def homeFax(self) -> dict:
        """
        Returns a dict of the data for the home fax. Returns None if no fields
        are set.

        Keys are "address_type", "email_address", "number",
        "original_display_name", and "original_entry_id".
        """
        try:
            return self._homeFax
        except AttributeError:
            data = {
                'address_type': self.homeFaxAddressType,
                'email_address': self.homeFaxEmailAddress,
                'number': self.homeFaxNumber,
                'original_display_name': self.homeFaxOriginalDisplayName,
                'original_entry_id': self.homeFaxOriginalEntryId,
            }
            self._homeFax = data if any(data[x] for x in data) else None
            return self._homeFax

    @property
    def homeFaxAddressType(self) -> str:
        """
        The type of address for the fax. MUST be set to "FAX" if present.
        """
        return self._ensureSetNamed('_businessFaxAddressType', '80D2')

    @property
    def homeFaxEmailAddress(self) -> str:
        """
        Contains a user-readable display name, followed by the "@" character,
        followed by a fax number.
        """
        return self._ensureSetNamed('_homeFaxAddressType', '80D3')

    @property
    def homeFaxNumber(self) -> str:
        """
        Contains the number of the contact's home fax.
        """
        return self._ensureSet('_homeFaxNumber', '__substg1.0_3A25')

    @property
    def homeFaxOriginalDisplayName(self) -> str:
        """
        The normalized subject for the contact.
        """
        return self._ensureSetNamed('_homeFaxAddressType', '80D4')

    @property
    def homeFaxOriginalEntryId(self) -> EntryID:
        """
        The one-off EntryID corresponding to this fax address.
        """
        return self._ensureSetNamed('_homeFaxAddressType', '80D5', overrideClass = EntryID.autoCreate)

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
    def lastModifiedBy(self) -> str:
        """
        The name of the last user to modify the contact file.
        """
        return self._ensureSet('_lastModifiedBy', '__substg1.0_3FFA')

    @property
    def mailAddress(self) -> str:
        """
        The complete mail address of the contact.
        """
        return self._ensureSet('_mailAddress', '__substg1.0_3A15')

    @property
    def mailAddressCountry(self) -> str:
        """
        The country portion of the contact's mail address.
        """
        return self._ensureSet('_mailAddressCountry', '__substg1.0_3A26')

    @property
    def mailAddressCountryCode(self) -> str:
        """
        The country code portion of the contact's mail address.
        """
        return self._ensureSetNamed('_mailAddressCountryCode', '80DD')

    @property
    def mailAddressLocality(self) -> str:
        """
        The locality or city portion of the contact's mail address.
        """
        return self._ensureSet('_mailAddressLocality', '__substg1.0_3A27')

    @property
    def mailAddressPostalCode(self) -> str:
        """
        The postal code portion of the contact's mail address.
        """
        return self._ensureSet('_mailAddressPostalCode', '__substg1.0_3A2A')

    @property
    def mailAddressPostOfficeBox(self) -> str:
        """
        The number or identifier of the contact's mail post office box.
        """
        return self._ensureSet('_mailAddressPostOfficeBox', '__substg1.0_3A2B')

    @property
    def mailAddressStateOrProvince(self) -> str:
        """
        The state or province portion of the contact's mail address.
        """
        return self._ensureSet('_mailAddressStateOrProvince', '__substg1.0_3A28')

    @property
    def mailAddressStreet(self) -> str:
        """
        The street portion of the contact's mail address.
        """
        return self._ensureSet('_mailAddressStreet', '__substg1.0_3A29')

    @property
    def middleNames(self) -> str:
        """
        The middle name(s) of the contact.
        """
        return self._ensureSet('_middleNames', '__substg1.0_3A44')

    @property
    def mobilePhone(self) -> str:
        """
        The mobile phone number of the contact.
        """
        return self._ensureSet('_mobilePhone', '__substg1.0_3A1C')

    @property
    def nickname(self) -> str:
        """
        The nickname of the contanct.
        """
        return self._ensureSet('_nickname', '__substg1.0_3A4F')

    @property
    def otherAddress(self) -> str:
        """
        The complete other address of the contact.
        """
        return self._ensureSetNamed('_otherAddress', '801C')

    @property
    def otherAddressCountry(self) -> str:
        """
        The country portion of the contact's other address.
        """
        return self._ensureSet('_otherAddressCountry', '__substg1.0_3A60')

    @property
    def otherAddressCountryCode(self) -> str:
        """
        The country code portion of the contact's other address.
        """
        return self._ensureSetNamed('_otherAddressCountryCode', '80DC')

    @property
    def otherAddressLocality(self) -> str:
        """
        The locality or city portion of the contact's other address.
        """
        return self._ensureSet('_otherAddressLocality', '__substg1.0_3A5F')

    @property
    def otherAddressPostalCode(self) -> str:
        """
        The postal code portion of the contact's other address.
        """
        return self._ensureSet('_otherAddressPostalCode', '__substg1.0_3A61')

    @property
    def otherAddressPostOfficeBox(self) -> str:
        """
        The number or identifier of the contact's other post office box.
        """
        return self._ensureSet('_otherAddressPostOfficeBox', '__substg1.0_3A64')

    @property
    def otherAddressStateOrProvince(self) -> str:
        """
        The state or province portion of the contact's other address.
        """
        return self._ensureSet('_otherAddressStateOrProvince', '__substg1.0_3A62')

    @property
    def otherAddressStreet(self) -> str:
        """
        The street portion of the contact's other address.
        """
        return self._ensureSet('_otherAddressStreet', '__substg1.0_3A63')

    @property
    def phoneticGivenName(self) -> str:
        """
        The phonetic pronunciation of the given name of the contact.
        """
        return self._ensureSetNamed('_phoneticGivenName', '802C')

    @property
    def phoneticSurname(self) -> str:
        """
        The phonetic pronunciation of the given name of the contact.
        """
        return self._ensureSetNamed('_phoneticSurname', '802D')

    @property
    def postalAddressID(self) -> PostalAddressID:
        """
        Indicates which physical address is the Mailing Address for this
        contact.
        """
        return self._ensureSetNamed('_postalAddressID', '8022', overrideClass = lambda x : PostalAddressID(x or 0), preserveNone = False)

    @property
    def primaryFax(self) -> dict:
        """
        Returns a dict of the data for the primary fax. Returns None if no
        fields are set.

        Keys are "address_type", "email_address", "number",
        "original_display_name", and "original_entry_id".
        """
        try:
            return self._primaryFax
        except AttributeError:
            data = {
                'address_type': self.primaryFaxAddressType,
                'email_address': self.primaryFaxEmailAddress,
                'number': self.primaryFaxNumber,
                'original_display_name': self.primaryFaxOriginalDisplayName,
                'original_entry_id': self.primaryFaxOriginalEntryId,
            }
            self._primaryFax = data if any(data[x] for x in data) else None
            return self._primaryFax

    @property
    def primaryFaxAddressType(self) -> str:
        """
        The type of address for the fax. MUST be set to "FAX" if present.
        """
        return self._ensureSetNamed('_primaryFaxAddressType', '80B2')

    @property
    def primaryFaxEmailAddress(self) -> str:
        """
        Contains a user-readable display name, followed by the "@" character,
        followed by a fax number.
        """
        return self._ensureSetNamed('_primaryFaxAddressType', '80B3')

    @property
    def primaryFaxNumber(self) -> str:
        """
        Contains the number of the contact's primary fax.
        """
        return self._ensureSet('_primaryFaxNumber', '__substg1.0_3A23')

    @property
    def primaryFaxOriginalDisplayName(self) -> str:
        """
        The normalized subject for the contact.
        """
        return self._ensureSetNamed('_primaryFaxAddressType', '80B4')

    @property
    def primaryFaxOriginalEntryId(self) -> EntryID:
        """
        The one-off EntryID corresponding to this fax address.
        """
        return self._ensureSetNamed('_primaryFaxAddressType', '80B5', overrideClass = EntryID.autoCreate)

    @property
    def spouseName(self) -> str:
        """
        The name of the contact's spouse.
        """
        return self._ensureSet('_spouseName', '__substg1.0_3A48')

    @property
    def surname(self) -> str:
        """
        The surname of the contact.
        """
        return self._ensureSet('_surname', '__substg1.0_3A11')

    @property
    def workAddress(self) -> str:
        """
        The complete work address of the contact.
        """
        return self._ensureSetNamed('_workAddress', '801B')

    @property
    def workAddressCountry(self) -> str:
        """
        The country portion of the contact's work address.
        """
        return self._ensureSetNamed('_workAddressCountry', '8049')

    @property
    def workAddressCountryCode(self) -> str:
        """
        The country code portion of the contact's work address.
        """
        return self._ensureSetNamed('_workAddressCountryCode', '80DB')

    @property
    def workAddressLocality(self) -> str:
        """
        The locality or city portion of the contact's work address.
        """
        return self._ensureSetNamed('_workAddressLocality', '8046')

    @property
    def workAddressPostalCode(self) -> str:
        """
        The postal code portion of the contact's work address.
        """
        return self._ensureSetNamed('_workAddressPostalCode', '8048')

    @property
    def workAddressPostOfficeBox(self) -> str:
        """
        The number or identifier of the contact's work post office box.
        """
        return self._ensureSetNamed('_workAddressPostOfficeBox', '804A')

    @property
    def workAddressStateOrProvince(self) -> str:
        """
        The state or province portion of the contact's work address.
        """
        return self._ensureSetNamed('_workAddressStateOrProvince', '8047')

    @property
    def workAddressStreet(self) -> str:
        """
        The street portion of the contact's work address.
        """
        return self._ensureSetNamed('_workAddressStreet', '8045')
