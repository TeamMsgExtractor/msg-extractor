import datetime

from typing import Dict, List, Set, Tuple, Union

from . import constants
from .enums import ContactLinkState, ElectronicAddressProperties, Gender, PostalAddressID
from .message_base import MessageBase
from .structures.entry_id import EntryID
from .structures.business_card import BusinessCardDisplayDefinition


class Contact(MessageBase):
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

    @property
    def account(self) -> str:
        """
        The account name of the contact.
        """
        return self._ensureSet('_account', '__substg1.0_3A00')

    @property
    def addressBookProviderArrayType(self) -> Set[ElectronicAddressProperties]:
        """
        A set of which Electronic Address properties are set on the contact.

        Property is stored in the MSG file as a sinlge int. The result should be
        identical to addressBookProviderEmailList.
        """
        return self._ensureSetNamed('_addressBookProviderArrayType', '8029', constants.PSETID_ADDRESS, ElectronicAddressProperties.fromBits)

    @property
    def addressBookProviderEmailList(self) -> Set[ElectronicAddressProperties]:
        """
        A set of which Electronic Address properties are set on the contact.
        """
        return self._ensureSetNamed('_addressBookProviderEmailList', '8028', constants.PSETID_ADDRESS, overrideClass = lambda x : {ElectronicAddressProperties(y) for y in x})

    @property
    def assistant(self) -> str:
        """
        The name of the contact's assistant.
        """
        return self._ensureSet('_assistant', '__substg1.0_3A30')

    @property
    def assistantTelephoneNumber(self) -> str:
        """
        Contains the telephone number of the contact's administrative assistant.
        """
        return self._ensureSet('_assistantTelephoneNumber', '__substg1.0_3A2E')

    @property
    def autoLog(self) -> bool:
        """
        Whether the client should create a Journal object for each action
        associated with the Contact object.
        """
        return self._ensureSetNamed('_autoLog', '8025', constants.PSETID_ADDRESS)

    @property
    def billing(self) -> str:
        """
        Billing information for the contact.
        """
        return self._ensureSetNamed('_billing', '8535', constants.PSETID_COMMON)

    @property
    def birthday(self) -> datetime.datetime:
        """
        The birthday of the contact at 11:59 UTC.
        """
        return self._ensureSetProperty('_birthday', '3A420040')

    @property
    def birthdayEventEntryID(self) -> EntryID:
        """
        The EntryID of an optional Appointement object that represents the
        contact's birtday.
        """
        return self._ensureSetNamed('_birthdayEventEntryID', '804D', constants.PSETID_ADDRESS, overrideClass = EntryID.autoCreate)

    @property
    def birthdayLocal(self) -> datetime.datetime:
        """
        The birthday of the contact at 0:00 in the client's local time zone.
        """
        return self._ensureSetNamed('_birthdayLocal', '80DE', constants.PSETID_ADDRESS)

    @property
    def businessCard(self) -> 'PIL.Image.Image':
        # First import PIL here so it's an optional dependency.
        try:
            import PIL.Image
            import PIL.ImageDraw
        except ImportError:
            raise ImportError('PIL or Pillow is required for generating the business card.') from None

        raise NotImplementedError('This function is not complete and as such is not yet functional. Please wait until later in the work cycle for this version.')

        # See "contact business card details.txt" for details.
        cardDef = self.businessCardDisplayDefinition
        # First things first, let's make the image we will be placing everything
        # onto.
        im = PIL.Image.new('RGB', (250, 150), cardDef.backgroundColor)
        # Create the ImageDraw instance so we can draw on it easily.
        imDraw = PIL.ImageDraw.ImageDraw(im)

        # Create the border of the image:
        imDraw.rectangle(((0, 0,), (249, 149)), outline = (109, 109, 109))


        # Finally, return the image.
        return im

    @property
    def businessCardCardPicture(self) -> bytes:
        """
        The image to be used on a business card. Must be either a PNG file or a
        JPEG file.
        """
        return self._ensureSetNamed('_businessCardCardPicture', '8041', constants.PSETID_ADDRESS)

    @property
    def businessCardDisplayDefinition(self) -> BusinessCardDisplayDefinition:
        """
        Specifies the customization details for displaying a contact as a
        business card.
        """
        return self._ensureSetNamed('_businessCardDisplayDefinition', '8040', constants.PSETID_ADDRESS, overrideClass = BusinessCardDisplayDefinition)

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
        return self._ensureSetNamed('_businessFaxAddressType', '80C2', constants.PSETID_ADDRESS)

    @property
    def businessFaxEmailAddress(self) -> str:
        """
        Contains a user-readable display name, followed by the "@" character,
        followed by a fax number.
        """
        return self._ensureSetNamed('_businessFaxEmailAddress', '80C3', constants.PSETID_ADDRESS)

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
        return self._ensureSetNamed('_businessFaxOriginalDisplayName', '80C4', constants.PSETID_ADDRESS)

    @property
    def businessFaxOriginalEntryId(self) -> EntryID:
        """
        The one-off EntryID corresponding to this fax address.
        """
        return self._ensureSetNamed('_businessFaxOriginalEntryId', '80C5', constants.PSETID_ADDRESS, overrideClass = EntryID.autoCreate)

    @property
    def businessTelephoneNumber(self) -> str:
        """
        Contains the number of the contact's business telephone.
        """
        return self._ensureSet('_businessTelephoneNumber', '__substg1.0_3A08')

    @property
    def businessTelephone2Number(self) -> Union[str, List[str]]:
        """
        Contains the second number or numbers of the contact's business.
        """
        return self._ensureSetTyped('_businessTelephone2Number', '3A1B')

    @property
    def businessHomePage(self) -> str:
        """
        Contains the url of the homepage of the contact's business.
        """
        return self._ensureSet('_businessHomePage', '__substg1.0_3A51')

    @property
    def callbackTelephoneNumber(self) -> str:
        """
        Contains the contact's callback telephone number.
        """
        return self._ensureSet('_callbackTelephoneNumber', '__substg1.0_3A02')

    @property
    def carTelephoneNumber(self) -> str:
        """
        Contains the number of the contact's car telephone.
        """
        return self._ensureSet('_carTelephoneNumber', '__substg1.0_3A1E')

    @property
    def childrensNames(self) -> List[str]:
        """
        A list of the named of the contact's children.
        """
        return self._ensureSetTyped('_childrensNames', '3A58')

    @property
    def companyMainTelephoneNumber(self) -> str:
        """
        Contains the number of the main telephone of the contact's company.
        """
        return self._ensureSet('_companyMainTelephoneNumber', '__substg1.0_3A57')

    @property
    def companyName(self) -> str:
        """
        The name of the company the contact works at.
        """
        return self._ensureSet('_companyName', '__substg1.0_3A16')

    @property
    def computerNetworkName(self) -> str:
        """
        The name of the network to wwhich the contact's computer is connected.
        """
        return self._ensureSet('_computerNetworkName', '__substg1.0_3A49')

    @property
    def contactCharacterSet(self) -> int:
        """
        The character set that is used for this Contact object.
        """
        return self._ensureSetNamed('_contactCharacterSet', '8023', constants.PSETID_ADDRESS)

    @property
    def contactItemData(self) -> List[int]:
        """
        Used to help display the contact information.
        """
        return self._ensureSetNamed('_contactItemData', '8007', constants.PSETID_ADDRESS)

    @property
    def contactLinkedGlobalAddressListEntryID(self) -> EntryID:
        """
        The EntryID of the GAL object to which the duplicate contact is linked.
        """
        return self._ensureSetNamed('_contactLinkedGlobalAddressListEntryID', '80E2', constants.PSETID_ADDRESS, overrideClass = EntryID.autoCreate)

    @property
    def contactLinkGlobalAddressListLinkID(self) -> str:
        """
        The GUID of the GAL contact to which the duplicate contact is linked.
        """
        return self._ensureSetNamed('_contactLinkGlobalAddressListLinkId', '80E8', constants.PSETID_ADDRESS)

    @property
    def contactLinkGlobalAddressListLinkState(self) -> ContactLinkState:
        """
        The state of linking between the GAL contact and the duplicate contact.
        """
        return self._ensureSetNamed('_contactLinkGlobalAddressListLinkState', '80E6', constants.PSETID_ADDRESS, overrideClass = ContactLinkState)

    @property
    def contactLinkLinkRejectHistory(self) -> List[bytes]:
        """
        A list of any contacts that were previously rejected for linking with
        the duplicate contact.
        """
        return self._ensureSetNamed('_contactLinkLinkRejectHistory', '80E5', constants.PSETID_ADDRESS)

    @property
    def contactLinkSMTPAddressCache(self) -> List[str]:
        """
        A list of the SMTP addresses that are used by the GAL contact that are
        linked to the duplicate contact.
        """
        return self._ensureSetNamed('_contactLinkSMTPAddressCache', '80E3', constants.PSETID_ADDRESS)

    @property
    def contactPhoto(self) -> bytes:
        """
        The contact photo, if it exists.
        """
        try:
            return self._contactPhoto
        except AttributeError:
            self._contactPhoto = None
            if self.hasPicture:
                if len(self.attachments) > 0:
                    contactPhotoAtt = next((att for att in self.attachments if att.isAttachmentContactPhoto), None)
                    if contactPhotoAtt:
                        self._contactPhoto = contactPhotoAtt.data
            return self._contactPhoto

    @property
    def contactUserField1(self) -> str:
        """
        Used to store custom text for a business card.
        """
        return self._ensureSetNamed('_contactUserField1', '804F', constants.PSETID_ADDRESS)

    @property
    def contactUserField2(self) -> str:
        """
        Used to store custom text for a business card.
        """
        return self._ensureSetNamed('_contactUserField2', '8050', constants.PSETID_ADDRESS)

    @property
    def contactUserField3(self) -> str:
        """
        Used to store custom text for a business card.
        """
        return self._ensureSetNamed('_contactUserField3', '8051', constants.PSETID_ADDRESS)

    @property
    def contactUserField4(self) -> str:
        """
        Used to store custom text for a business card.
        """
        return self._ensureSetNamed('_contactUserField4', '8052', constants.PSETID_ADDRESS)

    @property
    def customerID(self) -> str:
        """
        The contact's customer ID number.
        """
        return self._ensureSet('_customerID', '__substg1.0_3A4A')

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
        return self._ensureSetNamed('_email1AddressType', '8082', constants.PSETID_ADDRESS)

    @property
    def email1DisplayName(self) -> str:
        """
        The user-readable display name of the first email address.
        """
        return self._ensureSetNamed('_email1DisplayName', '8080', constants.PSETID_ADDRESS)

    @property
    def email1EmailAddress(self) -> str:
        """
        The first email address.
        """
        return self._ensureSetNamed('_email1EmailAddress', '8083', constants.PSETID_ADDRESS)

    @property
    def email1OriginalDisplayName(self) -> str:
        """
        The first SMTP email address that corresponds to the first email address
        for the contact.
        """
        return self._ensureSetNamed('_email1OriginalDisplayName', '8084', constants.PSETID_ADDRESS)

    @property
    def email1OriginalEntryId(self) -> EntryID:
        """
        The EntryID of the object correspinding to this electronic address.
        """
        return self._ensureSetNamed('_email1OriginalEntryId', '8085', constants.PSETID_ADDRESS, overrideClass = EntryID.autoCreate)

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
        return self._ensureSetNamed('_email2AddressType', '8092', constants.PSETID_ADDRESS)

    @property
    def email2DisplayName(self) -> str:
        """
        The user-readable display name of the second email address.
        """
        return self._ensureSetNamed('_email2DisplayName', '8090', constants.PSETID_ADDRESS)

    @property
    def email2EmailAddress(self) -> str:
        """
        The second email address.
        """
        return self._ensureSetNamed('_email2EmailAddress', '8093', constants.PSETID_ADDRESS)

    @property
    def email2OriginalDisplayName(self) -> str:
        """
        The second SMTP email address that corresponds to the second email address
        for the contact.
        """
        return self._ensureSetNamed('_email2OriginalDisplayName', '8094', constants.PSETID_ADDRESS)

    @property
    def email2OriginalEntryId(self) -> EntryID:
        """
        The EntryID of the object correspinding to this electronic address.
        """
        return self._ensureSetNamed('_email2OriginalEntryId', '8095', constants.PSETID_ADDRESS, overrideClass = EntryID.autoCreate)

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
        return self._ensureSetNamed('_email3AddressType', '80A2', constants.PSETID_ADDRESS)

    @property
    def email3DisplayName(self) -> str:
        """
        The user-readable display name of the third email address.
        """
        return self._ensureSetNamed('_email3DisplayName', '80A0', constants.PSETID_ADDRESS)

    @property
    def email3EmailAddress(self) -> str:
        """
        The third email address.
        """
        return self._ensureSetNamed('_email3EmailAddress', '80A3', constants.PSETID_ADDRESS)

    @property
    def email3OriginalDisplayName(self) -> str:
        """
        The third SMTP email address that corresponds to the third email address
        for the contact.
        """
        return self._ensureSetNamed('_email3OriginalDisplayName', '80A4', constants.PSETID_ADDRESS)

    @property
    def email3OriginalEntryId(self) -> EntryID:
        """
        The EntryID of the object correspinding to this electronic address.
        """
        return self._ensureSetNamed('_email3OriginalEntryId', '80A5', constants.PSETID_ADDRESS, overrideClass = EntryID.autoCreate)

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
        return self._ensureSetNamed('_fileUnder', '8005', constants.PSETID_ADDRESS)

    @property
    def fileUnderID(self) -> int:
        """
        The format to use for fileUnder. See PidLidFileUnderId in [MS-OXOCNTC]
        for details.
        """
        return self._ensureSetNamed('_fileUnderID', '8006', constants.PSETID_ADDRESS)

    @property
    def freeBusyLocation(self) -> str:
        """
        A URL path from which a client can retrieve free/busy status information
        for the contact as an iCalendat file.
        """
        return self._ensureSetNamed('_freeBusyLocation', '80D8', constants.PSETID_ADDRESS)

    @property
    def ftpSite(self) -> str:
        """
        The contact's File Transfer Protocol url.
        """
        return self._ensureSet('_ftpSite', '__substg1.0_3A4C')

    @property
    def gender(self) -> Gender:
        """
        The gender of the contact.
        """
        return self._ensureSet('_gender', '__substg1.0_3A4D', overrideClass = Gender)

    @property
    def generation(self) -> str:
        """
        A generational abbreviation that follows the full
        name of the contact.
        """
        return self._ensureSet('_generation', '__substg1.0_3A05')

    @property
    def givenName(self) -> str:
        """
        The first name of the contact.
        """
        return self._ensureSet('_givenName', '__substg1.0_3A06')

    @property
    def governmentIDNumber(self) -> str:
        """
        The contact's government ID number.
        """
        return self._ensureSet('_governmentIDNumber', '__substg1.0_3A07')

    @property
    def hasPicture(self) -> bool:
        """
        Whether the contact has a contact photo.
        """
        return self._ensureSetNamed('_hasPicture', '8015', constants.PSETID_ADDRESS, overrideClass = bool, preserveNone = False)

    @property
    def headerFormatProperties(self) -> constants.HEADER_FORMAT_TYPE:
        """
        Returns a dictionary of properties, in order, to be formatted into the
        header. Keys are the names to use in the header while the values are one
        of the following:
        None: Signifies no data was found for the property and it should be
            omitted from the header.
        str: A string to be formatted into the header using the string encoding.
        Tuple[Union[str, None], bool]: A string should be formatted into the
            header. If the bool is True, then place an empty string if the value
            is None, otherwise follow the same behavior as regular None.

        Additional note: If the value is an empty string, it will be dropped as
        well by default.
        """
        def strListToStr(inp : Union[str, List[str]]):
            """
            Small internal function for things that may return a string or list.
            """
            if isinstance(inp, str):
                return inp
            else:
                return ', '.join(inp)

        # Checking outlook printing, default behavior is to completely omit
        # *any* field that is not present. So while for extensability the
        # option exists to have it be present even if no data is found, we are
        # specifically not doing that.
        return {
            '-personal details-': {
                'Full Name': self.displayName,
                'Last Name': self.surname,
                'First Name': self.givenName,
                'Job Title': self.jobTitle,
                'Department': self.departmentName,
                'Company': self.companyName,
            },
            '-addresses-': {
                'Business Address': self.workAddress,
                'Home Address': self.homeAddress,
                'Other Address': self.otherAddress,
                'IM Address': self.instantMessagingAddress,
            },
            '-phone numbers-': {
                'Business:': self.businessTelephoneNumber,
                'Business 2': strListToStr(self.businessTelephone2Number),
                'Assistant': self.assistantTelephoneNumber,
                'Company Main Phone': self.companyMainTelephoneNumber,
                'Home': self.homeTelephoneNumber,
                'Home 2': strListToStr(self.homeTelephone2Number),
                'Mobile': self.mobileTelephoneNumber,
                'Car': self.carTelephoneNumber,
                'Radio': self.radioTelephoneNumber,
                'Pager': self.pagerTelephoneNumber,
                'Callback': self.callbackTelephoneNumber,
                'Other': self.otherTelephoneNumber,
                'Primary Phone': self.primaryTelephoneNumber,
                'Telex': self.telexNumber,
                'TTY/TDD Phone': self.tddTelephoneNumber,
                'ISDN': self.isdnNumber,
                'Business Fax': self.businessFaxNumber,
                'Home Fax': self.homeFaxNumber,
            },
            '-emails-': {
                'Email': self.email1EmailAddress or self.email1OriginalDisplayName,
                'Email Display As': self.email1DisplayName,
                'Email 2': self.email2EmailAddress or self.email2OriginalDisplayName,
                'Email2 Display As': self.email2DisplayName,
                'Email 3': self.email3EmailAddress or self.email3OriginalDisplayName,
                'Email3 Display As': self.email3DisplayName,
            },
            '-other-': {
                'Birthday': self.birthday.__format__('%B %d, %Y') if self.birthdayLocal else None,
                'Anniversary': self.weddingAnniversary.__format__('%B %d, %Y') if self.weddingAnniversaryLocal else None,
                'Spouse/Partner': self.spouseName,
                'Profession': self.profession,
                'Children': ', '.join(self.childrensNames),
                'Hobbies': self.hobbies,
                'Assistant': self.assistant,
                'Web Page': self.webpageUrl,
            },
        }

    @property
    def hobbies(self) -> str:
        """
        The hobies of the contact.
        """
        return self._ensureSet('_hobbies', '__substg1.0_3A43')

    @property
    def homeAddress(self) -> str:
        """
        The complete home address of the contact.
        """
        return self._ensureSetNamed('_homeAddress', '801A', constants.PSETID_ADDRESS)

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
        return self._ensureSetNamed('_homeAddressCountryCode', '80DA', constants.PSETID_ADDRESS)

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
        return self._ensureSetNamed('_homeFaxAddressType', '80D2', constants.PSETID_ADDRESS)

    @property
    def homeFaxEmailAddress(self) -> str:
        """
        Contains a user-readable display name, followed by the "@" character,
        followed by a fax number.
        """
        return self._ensureSetNamed('_homeFaxEmailAddress', '80D3', constants.PSETID_ADDRESS)

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
        return self._ensureSetNamed('_homeFaxOriginalDisplayName', '80D4', constants.PSETID_ADDRESS)

    @property
    def homeFaxOriginalEntryId(self) -> EntryID:
        """
        The one-off EntryID corresponding to this fax address.
        """
        return self._ensureSetNamed('_homeFaxOriginalEntryId', '80D5', constants.PSETID_ADDRESS, overrideClass = EntryID.autoCreate)

    @property
    def homeTelephoneNumber(self) -> str:
        """
        The number of the contact's home telephone.
        """
        return self._ensureSet('_homeTelephoneNumber', '__substg1.0_3A09')

    @property
    def homeTelephone2Number(self) -> Union[str, List[str]]:
        """
        The number(s) of the contact's second home telephone.
        """
        return self._ensureSetTyped('_homeTelephone2Number', '3A2F')

    @property
    def initials(self) -> str:
        """
        The initials of the contact.
        """
        return self._ensureSet('_initials', '__substg1.0_3A0A')

    @property
    def instantMessagingAddress(self) -> str:
        """
        The instant messaging address of the contact.
        """
        return self._ensureSetNamed('_instantMessagingAddress', '8062', constants.PSETID_ADDRESS)

    @property
    def isContactLinked(self) -> bool:
        """
        Whether the contact is linked to other contacts.
        """
        return self._ensureSetNamed('_isContactLinked', '80E0', constants.PSETID_ADDRESS)

    @property
    def isdnNumber(self) -> str:
        """
        The Integrated Services Digital Network (ISDN) telephone number of the
        contact.
        """
        return self._ensureSet('_isdnNumber', '__substg1.0_3A2D')

    @property
    def jobTitle(self) -> str:
        """
        The job title of the contact.
        """
        return self._ensureSet('_jobTitle', '__substg1.0_3A17')

    @property
    def language(self) -> str:
        """
        The language that the contact uses.
        """
        return self._ensureSet('_language', '__substg1.0_3A0C')

    @property
    def lastModifiedBy(self) -> str:
        """
        The name of the last user to modify the contact file.
        """
        return self._ensureSet('_lastModifiedBy', '__substg1.0_3FFA')

    @property
    def location(self) -> str:
        """
        The location of the contact. For example, this could be the building or
        office number of the contact.
        """
        return self._ensureSet('_location', '__substg1.0_3A0D')

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
        return self._ensureSetNamed('_mailAddressCountryCode', '80DD', constants.PSETID_ADDRESS)

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
    def managerName(self) -> str:
        """
        The name of the contact's manager.
        """
        return self._ensureSet('_managerName', '__substg1.0_3A4E')

    @property
    def middleName(self) -> str:
        """
        The middle name(s) of the contact.
        """
        return self._ensureSet('_middleNames', '__substg1.0_3A44')

    @property
    def mobileTelephoneNumber(self) -> str:
        """
        The mobile telephone number of the contact.
        """
        return self._ensureSet('_mobileTelephoneNumber', '__substg1.0_3A1C')

    @property
    def nickname(self) -> str:
        """
        The nickname of the contanct.
        """
        return self._ensureSet('_nickname', '__substg1.0_3A4F')

    @property
    def officeLocation(self) -> str:
        """
        The location of the office that the contact works in.
        """
        return self._ensureSet('_officeLocation', '__substg1.0_3A19')

    @property
    def organizationalIDNumber(self) -> str:
        """
        The organizational ID number for the contact, such as an employee ID
        number.
        """
        return self._ensureSet('_organizationalIdNumber', '__substg1.0_3A10')

    @property
    def oscSyncEnabled(self) -> bool:
        """
        Whether contact synchronization with an external source (such as a
        social networking site) is handled by the server.
        """
        return self._ensureSetProperty('_oscSyncEnabled', '7C24000B')

    @property
    def otherAddress(self) -> str:
        """
        The complete other address of the contact.
        """
        return self._ensureSetNamed('_otherAddress', '801C', constants.PSETID_ADDRESS)

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
        return self._ensureSetNamed('_otherAddressCountryCode', '80DC', constants.PSETID_ADDRESS)

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
    def otherTelephoneNumber(self) -> str:
        """
        Contains the number of the contact's other telephone.
        """
        return self._ensureSet('_otherTelephoneNumber', '__substg1.0_3A1F')

    @property
    def pagerTelephoneNumber(self) -> str:
        """
        The contact's pager telephone number.
        """
        return self._ensureSet('_pagerTelephoneNumber', '__substg1.0_3A21')

    @property
    def personalHomePage(self) -> str:
        """
        The contact's personal web page UL.
        """
        return self._ensureSet('_personalHomePage', '__substg1.0_3A50')

    @property
    def phoneticCompanyName(self) -> str:
        """
        The phonetic pronunciation of the contact's company name.
        """
        return self._ensureSetNamed('_phoneticCompanyName', '802E', constants.PSETID_ADDRESS)

    @property
    def phoneticGivenName(self) -> str:
        """
        The phonetic pronunciation of the contact's given name.
        """
        return self._ensureSetNamed('_phoneticGivenName', '802C', constants.PSETID_ADDRESS)

    @property
    def phoneticSurname(self) -> str:
        """
        The phonetic pronunciation of the given name of the contact.
        """
        return self._ensureSetNamed('_phoneticSurname', '802D', constants.PSETID_ADDRESS)

    @property
    def postalAddressID(self) -> PostalAddressID:
        """
        Indicates which physical address is the Mailing Address for this
        contact.
        """
        return self._ensureSetNamed('_postalAddressID', '8022', constants.PSETID_ADDRESS, overrideClass = lambda x : PostalAddressID(x or 0), preserveNone = False)

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
        return self._ensureSetNamed('_primaryFaxAddressType', '80B2', constants.PSETID_ADDRESS)

    @property
    def primaryFaxEmailAddress(self) -> str:
        """
        Contains a user-readable display name, followed by the "@" character,
        followed by a fax number.
        """
        return self._ensureSetNamed('_primaryFaxEmailAddress', '80B3', constants.PSETID_ADDRESS)

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
        return self._ensureSetNamed('_primaryFaxOriginalDisplayName', '80B4', constants.PSETID_ADDRESS)

    @property
    def primaryFaxOriginalEntryId(self) -> EntryID:
        """
        The one-off EntryID corresponding to this fax address.
        """
        return self._ensureSetNamed('_primaryFaxOriginalEntryId', '80B5', constants.PSETID_ADDRESS, overrideClass = EntryID.autoCreate)

    @property
    def primaryTelephoneNumber(self) -> str:
        """
        Contains the number of the contact's primary telephone.
        """
        return self._ensureSet('_primaryTelephoneNumber', '__substg1.0_3A1A')

    @property
    def profession(self) -> str:
        """
        The profession of the contact.
        """
        return self._ensureSet('_profession', '__substg1.0_3A46')

    @property
    def radioTelephoneNumber(self) -> str:
        """
        Contains the number of the contact's radio telephone.
        """
        return self._ensureSet('_radioTelephoneNumber', '__substg1.0_3A1D')

    @property
    def referenceEntryID(self) -> EntryID:
        """
        Contains a value that is equal to the value of the EntryID of the
        Contact object unless the Contact object is a copy of an earlier
        original.
        """
        return self._ensureSetNamed('_referenceEntryID', '85BD', constants.PSETID_COMMON, overrideClass = EntryID.autoCreate)

    @property
    def referredByName(self) -> str:
        """
        The name of the person who referred this contact to the user.
        """
        return self._ensureSet('_referredByName', '__substg1.0_3A47')

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
    def tddTelephoneNumber(self) -> str:
        """
        The telephone number for the contact's text telephone (TTY) or
        telecommunication device for the deaf (TDD).
        """
        return self._ensureSet('_tddTelephoneNumber', '__substg1.0_3A4B')

    @property
    def telexNumber(self) -> Union[str, List[str]]:
        """
        The contact's telex number(s).
        """
        return self._ensureSetTyped('_telexNumber', '3A2C')

    @property
    def userX509Certificate(self) -> List[bytes]:
        """
        A list of certificates for the contact.
        """
        return self._ensureSetTyped('_userX509Certificate', '3A70')

    @property
    def weddingAnniversary(self) -> datetime.datetime:
        """
        The wedding anniversary of the contact at 11:59 UTC.
        """
        return self._ensureSetProperty('_weddingAnniversary', '3A410040')

    @property
    def weddingAnniversaryEventEntryID(self) -> EntryID:
        """
        The EntryID of an optional Appointement object that represents the
        contact's wedding anniversary.
        """
        return self._ensureSetNamed('_weddingAnniversaryEventEntryID', '804E', constants.PSETID_ADDRESS, overrideClass = EntryID.autoCreate)

    @property
    def weddingAnniversaryLocal(self) -> datetime.datetime:
        """
        The wedding anniversary of the contact at 0:00 in the client's local
        time zone.
        """
        return self._ensureSetNamed('_weddingAnniversaryLocal', '80DF', constants.PSETID_ADDRESS)

    @property
    def webpageUrl(self) -> str:
        """
        The contact's business web page url. SHOULD be the same as businessUrl.
        """
        return self._ensureSetNamed('_webpageUrl', '802B', constants.PSETID_ADDRESS)

    @property
    def workAddress(self) -> str:
        """
        The complete work address of the contact.
        """
        return self._ensureSetNamed('_workAddress', '801B', constants.PSETID_ADDRESS)

    @property
    def workAddressCountry(self) -> str:
        """
        The country portion of the contact's work address.
        """
        return self._ensureSetNamed('_workAddressCountry', '8049', constants.PSETID_ADDRESS)

    @property
    def workAddressCountryCode(self) -> str:
        """
        The country code portion of the contact's work address.
        """
        return self._ensureSetNamed('_workAddressCountryCode', '80DB', constants.PSETID_ADDRESS)

    @property
    def workAddressLocality(self) -> str:
        """
        The locality or city portion of the contact's work address.
        """
        return self._ensureSetNamed('_workAddressLocality', '8046', constants.PSETID_ADDRESS)

    @property
    def workAddressPostalCode(self) -> str:
        """
        The postal code portion of the contact's work address.
        """
        return self._ensureSetNamed('_workAddressPostalCode', '8048', constants.PSETID_ADDRESS)

    @property
    def workAddressPostOfficeBox(self) -> str:
        """
        The number or identifier of the contact's work post office box.
        """
        return self._ensureSetNamed('_workAddressPostOfficeBox', '804A', constants.PSETID_ADDRESS)

    @property
    def workAddressStateOrProvince(self) -> str:
        """
        The state or province portion of the contact's work address.
        """
        return self._ensureSetNamed('_workAddressStateOrProvince', '8047', constants.PSETID_ADDRESS)

    @property
    def workAddressStreet(self) -> str:
        """
        The street portion of the contact's work address.
        """
        return self._ensureSetNamed('_workAddressStreet', '8045', constants.PSETID_ADDRESS)
