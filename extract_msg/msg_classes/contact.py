__all__ = [
    'Contact',
]


import datetime
import functools
import io
import json

from typing import Dict, List, Optional, Set, Tuple, Union

from ..constants import HEADER_FORMAT_TYPE, ps
from ..enums import (
        ContactLinkState, ElectronicAddressProperties, Gender,
        InsecureFeatures, PostalAddressID
    )
from ..exceptions import SecurityError
from .message_base import MessageBase
from ..structures.business_card import BusinessCardDisplayDefinition
from ..structures.entry_id import EntryID


# I would use TypeAlias, but my type checker *really* hates that.
_EMAIL_DICT = Dict[str, Optional[Union[str, EntryID]]]
_FAX_DICT = _EMAIL_DICT


class Contact(MessageBase):
    """
    Class used for parsing contacts.
    """

    def getJson(self) -> str:
        # To save a lot of trouble and repetiion, just return a JSON version of
        # the header format properties.
        return json.dumps(self.headerFormatProperties)

    @functools.cached_property
    def account(self) -> Optional[str]:
        """
        The account name of the contact.
        """
        return self.getStringStream('__substg1.0_3A00')

    @functools.cached_property
    def addressBookProviderArrayType(self) -> Optional[ElectronicAddressProperties]:
        """
        A union of which Electronic Address properties are set on the contact.

        Property is stored in the MSG file as a sinlge int. The result should be
        a union of the flags specified by addressBookProviderEmailList.
        """
        return self.getNamedAs('8029', ps.PSETID_ADDRESS, ElectronicAddressProperties)

    @functools.cached_property
    def addressBookProviderEmailList(self) -> Optional[Set[ElectronicAddressProperties]]:
        """
        A set of which Electronic Address properties are set on the contact.
        """
        return self.getNamedAs('8028', ps.PSETID_ADDRESS, ElectronicAddressProperties.fromIter)

    @functools.cached_property
    def assistant(self) -> Optional[str]:
        """
        The name of the contact's assistant.
        """
        return self.getStringStream('__substg1.0_3A30')

    @functools.cached_property
    def assistantTelephoneNumber(self) -> Optional[str]:
        """
        Contains the telephone number of the contact's administrative assistant.
        """
        return self.getStringStream('__substg1.0_3A2E')

    @functools.cached_property
    def autoLog(self) -> bool:
        """
        Whether the client should create a Journal object for each action
        associated with the Contact object.
        """
        return bool(self.getNamedProp('8025', ps.PSETID_ADDRESS))

    @functools.cached_property
    def billing(self) -> Optional[str]:
        """
        Billing information for the contact.
        """
        return self.getNamedProp('8535', ps.PSETID_COMMON)

    @functools.cached_property
    def birthday(self) -> Optional[datetime.datetime]:
        """
        The birthday of the contact at 11:59 UTC.
        """
        return self.getPropertyVal('3A420040')

    @functools.cached_property
    def birthdayEventEntryID(self) -> Optional[EntryID]:
        """
        The EntryID of an optional Appointement object that represents the
        contact's birtday.
        """
        return self.getNamedAs('804D', ps.PSETID_ADDRESS, EntryID.autoCreate)

    @functools.cached_property
    def birthdayLocal(self) -> Optional[datetime.datetime]:
        """
        The birthday of the contact at 0:00 in the client's local time zone.
        """
        return self.getNamedProp('80DE', ps.PSETID_ADDRESS)

    @functools.cached_property
    def businessCard(self) -> bytes:
        if InsecureFeatures.PIL_IMAGE_PARSING not in self.insecureFeatures:
            return SecurityError('PIL_IMAGE_PARSING must be enabled to create a business card image.')
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
        imDraw.rectangle(((0, 0), (249, 149)), outline = (109, 109, 109))


        # Finally, return the image.
        out = io.BytesIO()
        im.save(out, 'png')
        return out

    @functools.cached_property
    def businessCardCardPicture(self) -> Optional[bytes]:
        """
        The image to be used on a business card.

        Must be either a PNG file or a JPEG file.
        """
        return self.getNamedProp('8041', ps.PSETID_ADDRESS)

    @functools.cached_property
    def businessCardDisplayDefinition(self) -> Optional[BusinessCardDisplayDefinition]:
        """
        Specifies the customization details for displaying a contact as a
        business card.
        """
        return self.getNamedAs('8040', ps.PSETID_ADDRESS, BusinessCardDisplayDefinition)

    @functools.cached_property
    def businessFax(self) -> Optional[_FAX_DICT]:
        """
        Returns a ``dict`` of the data for the business fax.

        Returns ``None`` if no fields are set.

        Keys are "address_type", "email_address", "number",
        "original_display_name", and "original_entry_id".
        """
        data = {
            'address_type': self.businessFaxAddressType,
            'email_address': self.businessFaxEmailAddress,
            'number': self.businessFaxNumber,
            'original_display_name': self.businessFaxOriginalDisplayName,
            'original_entry_id': self.businessFaxOriginalEntryId,
        }
        return data if any(data[x] for x in data) else None

    @functools.cached_property
    def businessFaxAddressType(self) -> Optional[str]:
        """
        The type of address for the fax.

        MUST be set to "FAX" if present.
        """
        return self.getNamedProp('80C2', ps.PSETID_ADDRESS)

    @functools.cached_property
    def businessFaxEmailAddress(self) -> Optional[str]:
        """
        Contains a user-readable display name, followed by the "@" character,
        followed by a fax number.
        """
        return self.getNamedProp('80C3', ps.PSETID_ADDRESS)

    @functools.cached_property
    def businessFaxNumber(self) -> Optional[str]:
        """
        Contains the number of the contact's business fax.
        """
        return self.getStringStream('__substg1.0_3A24')

    @functools.cached_property
    def businessFaxOriginalDisplayName(self) -> Optional[str]:
        """
        The normalized subject for the contact.
        """
        return self.getNamedProp('80C4', ps.PSETID_ADDRESS)

    @functools.cached_property
    def businessFaxOriginalEntryId(self) -> Optional[EntryID]:
        """
        The one-off EntryID corresponding to this fax address.
        """
        return self.getNamedAs('80C5', ps.PSETID_ADDRESS, EntryID.autoCreate)

    @functools.cached_property
    def businessTelephoneNumber(self) -> Optional[str]:
        """
        Contains the number of the contact's business telephone.
        """
        return self.getStringStream('__substg1.0_3A08')

    @functools.cached_property
    def businessTelephone2Number(self) -> Optional[Union[str, List[str]]]:
        """
        Contains the second number or numbers of the contact's business.
        """
        return self.getSingleOrMultipleString('__substg1.0_3A1B')

    @functools.cached_property
    def businessHomePage(self) -> Optional[str]:
        """
        Contains the url of the homepage of the contact's business.
        """
        return self.getStringStream('__substg1.0_3A51')

    @functools.cached_property
    def callbackTelephoneNumber(self) -> Optional[str]:
        """
        Contains the contact's callback telephone number.
        """
        return self.getStringStream('__substg1.0_3A02')

    @functools.cached_property
    def carTelephoneNumber(self) -> Optional[str]:
        """
        Contains the number of the contact's car telephone.
        """
        return self.getStringStream('__substg1.0_3A1E')

    @functools.cached_property
    def childrensNames(self) -> Optional[List[str]]:
        """
        A list of the named of the contact's children.
        """
        return self.getMultipleString('__substg1.0_3A58')

    @functools.cached_property
    def companyMainTelephoneNumber(self) -> Optional[str]:
        """
        Contains the number of the main telephone of the contact's company.
        """
        return self.getStringStream('__substg1.0_3A57')

    @functools.cached_property
    def companyName(self) -> Optional[str]:
        """
        The name of the company the contact works at.
        """
        return self.getStringStream('__substg1.0_3A16')

    @functools.cached_property
    def computerNetworkName(self) -> Optional[str]:
        """
        The name of the network to wwhich the contact's computer is connected.
        """
        return self.getStringStream('__substg1.0_3A49')

    @functools.cached_property
    def contactCharacterSet(self) -> Optional[int]:
        """
        The character set that is used for this Contact object.
        """
        return self.getNamedProp('8023', ps.PSETID_ADDRESS)

    @functools.cached_property
    def contactItemData(self) -> Optional[List[int]]:
        """
        Used to help display the contact information.
        """
        return self.getNamedProp('8007', ps.PSETID_ADDRESS)

    @functools.cached_property
    def contactLinkedGlobalAddressListEntryID(self) -> Optional[EntryID]:
        """
        The EntryID of the GAL object to which the duplicate contact is linked.
        """
        return self.getNamedAs('80E2', ps.PSETID_ADDRESS, EntryID.autoCreate)

    @functools.cached_property
    def contactLinkGlobalAddressListLinkID(self) -> Optional[str]:
        """
        The GUID of the GAL contact to which the duplicate contact is linked.
        """
        return self.getNamedProp('80E8', ps.PSETID_ADDRESS)

    @functools.cached_property
    def contactLinkGlobalAddressListLinkState(self) -> Optional[ContactLinkState]:
        """
        The state of linking between the GAL contact and the duplicate contact.
        """
        return self.getNamedAs('80E6', ps.PSETID_ADDRESS, ContactLinkState)

    @functools.cached_property
    def contactLinkLinkRejectHistory(self) -> Optional[List[bytes]]:
        """
        A list of any contacts that were previously rejected for linking with
        the duplicate contact.
        """
        return self.getNamedProp('80E5', ps.PSETID_ADDRESS)

    @functools.cached_property
    def contactLinkSMTPAddressCache(self) -> Optional[List[str]]:
        """
        A list of the SMTP addresses that are used by the GAL contact that are
        linked to the duplicate contact.
        """
        return self.getNamedProp('80E3', ps.PSETID_ADDRESS)

    @functools.cached_property
    def contactPhoto(self) -> Optional[bytes]:
        """
        The contact photo, if it exists.
        """
        if self.hasPicture:
            if len(self.attachments) > 0:
                contactPhotoAtt = next((att for att in self.attachments if att.isAttachmentContactPhoto), None)
                if contactPhotoAtt:
                    return contactPhotoAtt.data
        return None

    @functools.cached_property
    def contactUserField1(self) -> Optional[str]:
        """
        Used to store custom text for a business card.
        """
        return self.getNamedProp('804F', ps.PSETID_ADDRESS)

    @functools.cached_property
    def contactUserField2(self) -> Optional[str]:
        """
        Used to store custom text for a business card.
        """
        return self.getNamedProp('8050', ps.PSETID_ADDRESS)

    @functools.cached_property
    def contactUserField3(self) -> Optional[str]:
        """
        Used to store custom text for a business card.
        """
        return self.getNamedProp('8051', ps.PSETID_ADDRESS)

    @functools.cached_property
    def contactUserField4(self) -> Optional[str]:
        """
        Used to store custom text for a business card.
        """
        return self.getNamedProp('8052', ps.PSETID_ADDRESS)

    @functools.cached_property
    def customerID(self) -> Optional[str]:
        """
        The contact's customer ID number.
        """
        return self.getStringStream('__substg1.0_3A4A')

    @functools.cached_property
    def departmentName(self) -> Optional[str]:
        """
        The name of the department the contact works in.
        """
        return self.getStringStream('__substg1.0_3A18')

    @functools.cached_property
    def displayName(self) -> Optional[str]:
        """
        The full name of the contact.
        """
        return self.getStringStream('__substg1.0_3001')

    @functools.cached_property
    def displayNamePrefix(self) -> Optional[str]:
        """
        The contact's honorific title.
        """
        return self.getStringStream('__substg1.0_3A45')

    @functools.cached_property
    def email1(self) -> Optional[_EMAIL_DICT]:
        """
        Returns a ``dict`` of the data for email 1.

        Returns ``None`` if no fields are set.

        Keys are "address_type", "display_name", "email_address",
        "original_display_name", and "original_entry_id".
        """
        data = {
            'address_type': self.email1AddressType,
            'display_name': self.email1DisplayName,
            'email_address': self.email1EmailAddress,
            'original_display_name': self.email1OriginalDisplayName,
            'original_entry_id': self.email1OriginalEntryId,
        }
        return data if any(data[x] for x in data) else None

    @functools.cached_property
    def email1AddressType(self) -> Optional[str]:
        """
        The address type of the first email address.
        """
        return self.getNamedProp('8082', ps.PSETID_ADDRESS)

    @functools.cached_property
    def email1DisplayName(self) -> Optional[str]:
        """
        The user-readable display name of the first email address.
        """
        return self.getNamedProp('8080', ps.PSETID_ADDRESS)

    @functools.cached_property
    def email1EmailAddress(self) -> Optional[str]:
        """
        The first email address.
        """
        return self.getNamedProp('8083', ps.PSETID_ADDRESS)

    @functools.cached_property
    def email1OriginalDisplayName(self) -> Optional[str]:
        """
        The first SMTP email address that corresponds to the first email address
        for the contact.
        """
        return self.getNamedProp('8084', ps.PSETID_ADDRESS)

    @functools.cached_property
    def email1OriginalEntryId(self) -> Optional[EntryID]:
        """
        The EntryID of the object correspinding to this electronic address.
        """
        return self.getNamedAs('8085', ps.PSETID_ADDRESS, EntryID.autoCreate)

    @functools.cached_property
    def email2(self) -> Optional[_EMAIL_DICT]:
        """
        Returns a ``dict`` of the data for email 2.

        Returns ``None`` if no fields are set.
        """
        data = {
            'address_type': self.email2AddressType,
            'display_name': self.email2DisplayName,
            'email_address': self.email2EmailAddress,
            'original_display_name': self.email2OriginalDisplayName,
            'original_entry_id': self.email2OriginalEntryId,
        }
        return data if any(data[x] for x in data) else None

    @functools.cached_property
    def email2AddressType(self) -> Optional[str]:
        """
        The address type of the second email address.
        """
        return self.getNamedProp('8092', ps.PSETID_ADDRESS)

    @functools.cached_property
    def email2DisplayName(self) -> Optional[str]:
        """
        The user-readable display name of the second email address.
        """
        return self.getNamedProp('8090', ps.PSETID_ADDRESS)

    @functools.cached_property
    def email2EmailAddress(self) -> Optional[str]:
        """
        The second email address.
        """
        return self.getNamedProp('8093', ps.PSETID_ADDRESS)

    @functools.cached_property
    def email2OriginalDisplayName(self) -> Optional[str]:
        """
        The second SMTP email address that corresponds to the second email
        address for the contact.
        """
        return self.getNamedProp('8094', ps.PSETID_ADDRESS)

    @functools.cached_property
    def email2OriginalEntryId(self) -> Optional[EntryID]:
        """
        The EntryID of the object correspinding to this electronic address.
        """
        return self.getNamedAs('8095', ps.PSETID_ADDRESS, EntryID.autoCreate)

    @functools.cached_property
    def email3(self) -> Optional[_EMAIL_DICT]:
        """
        Returns a ``dict`` of the data for email 3.

        Returns ``None`` if no fields are set.
        """
        data = {
            'address_type': self.email3AddressType,
            'display_name': self.email3DisplayName,
            'email_address': self.email3EmailAddress,
            'original_display_name': self.email3OriginalDisplayName,
            'original_entry_id': self.email3OriginalEntryId,
        }
        return data if any(data[x] for x in data) else None

    @functools.cached_property
    def email3AddressType(self) -> Optional[str]:
        """
        The address type of the third email address.
        """
        return self.getNamedProp('80A2', ps.PSETID_ADDRESS)

    @functools.cached_property
    def email3DisplayName(self) -> Optional[str]:
        """
        The user-readable display name of the third email address.
        """
        return self.getNamedProp('80A0', ps.PSETID_ADDRESS)

    @functools.cached_property
    def email3EmailAddress(self) -> Optional[str]:
        """
        The third email address.
        """
        return self.getNamedProp('80A3', ps.PSETID_ADDRESS)

    @functools.cached_property
    def email3OriginalDisplayName(self) -> Optional[str]:
        """
        The third SMTP email address that corresponds to the third email address
        for the contact.
        """
        return self.getNamedProp('80A4', ps.PSETID_ADDRESS)

    @functools.cached_property
    def email3OriginalEntryId(self) -> Optional[EntryID]:
        """
        The EntryID of the object correspinding to this electronic address.
        """
        return self.getNamedAs('80A5', ps.PSETID_ADDRESS, EntryID.autoCreate)

    @functools.cached_property
    def emails(self) -> Tuple[Union[_EMAIL_DICT, None], Union[_EMAIL_DICT, None], Union[_EMAIL_DICT, None]]:
        """
        Returns a tuple of all the email ``dict``\\s.

        Value for an email will be ``None`` if no fields were set.
        """
        return (self.email1, self.email2, self.email3)

    @functools.cached_property
    def faxNumbers(self) -> Dict[str, Optional[_FAX_DICT]]:
        """
        Returns a ``dict`` of the fax numbers.

        Entry will be ``None`` if no fields were set.

        Keys are "business", "home", and "primary".
        """
        return {
            'business': self.businessFax,
            'home': self.homeFax,
            'primary': self.primaryFax,
        }

    @functools.cached_property
    def fileUnder(self) -> Optional[str]:
        """
        The name under which to file a contact when displaying a list of
        contacts.
        """
        return self.getNamedProp('8005', ps.PSETID_ADDRESS)

    @functools.cached_property
    def fileUnderID(self) -> Optional[int]:
        """
        The format to use for fileUnder. See PidLidFileUnderId in [MS-OXOCNTC]
        for details.
        """
        return self.getNamedProp('8006', ps.PSETID_ADDRESS)

    @functools.cached_property
    def freeBusyLocation(self) -> Optional[str]:
        """
        A URL path from which a client can retrieve free/busy status information
        for the contact as an iCalendat file.
        """
        return self.getNamedProp('80D8', ps.PSETID_ADDRESS)

    @functools.cached_property
    def ftpSite(self) -> Optional[str]:
        """
        The contact's File Transfer Protocol url.
        """
        return self.getStringStream('__substg1.0_3A4C')

    @functools.cached_property
    def gender(self) -> Gender:
        """
        The gender of the contact.
        """
        return Gender(self.getPropertyVal('3A4D0002', 0))

    @functools.cached_property
    def generation(self) -> Optional[str]:
        """
        A generational abbreviation that follows the full name of the contact.
        """
        return self.getStringStream('__substg1.0_3A05')

    @functools.cached_property
    def givenName(self) -> Optional[str]:
        """
        The first name of the contact.
        """
        return self.getStringStream('__substg1.0_3A06')

    @functools.cached_property
    def governmentIDNumber(self) -> Optional[str]:
        """
        The contact's government ID number.
        """
        return self.getStringStream('__substg1.0_3A07')

    @functools.cached_property
    def hasPicture(self) -> bool:
        """
        Whether the contact has a contact photo.
        """
        return bool(self.getNamedProp('8015', ps.PSETID_ADDRESS))

    @property
    def headerFormatProperties(self) -> HEADER_FORMAT_TYPE:
        def strListToStr(inp: Optional[Union[str, List[str]]]):
            """
            Small internal function for things that may return a string or list.
            """
            if inp is None or isinstance(inp, str):
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
                'Business': self.businessTelephoneNumber,
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
                'Birthday': self.birthday.__format__(self.dateFormat) if self.birthdayLocal else None,
                'Anniversary': self.weddingAnniversary.__format__(self.dateFormat) if self.weddingAnniversaryLocal else None,
                'Spouse/Partner': self.spouseName,
                'Profession': self.profession,
                'Children': strListToStr(self.childrensNames),
                'Hobbies': self.hobbies,
                'Assistant': self.assistant,
                'Web Page': self.webpageUrl,
            },
        }

    @functools.cached_property
    def hobbies(self) -> Optional[str]:
        """
        The hobies of the contact.
        """
        return self.getStringStream('__substg1.0_3A43')

    @functools.cached_property
    def homeAddress(self) -> Optional[str]:
        """
        The complete home address of the contact.
        """
        return self.getNamedProp('801A', ps.PSETID_ADDRESS)

    @functools.cached_property
    def homeAddressCountry(self) -> Optional[str]:
        """
        The country portion of the contact's home address.
        """
        return self.getStringStream('__substg1.0_3A5A')

    @functools.cached_property
    def homeAddressCountryCode(self) -> Optional[str]:
        """
        The country code portion of the contact's home address.
        """
        return self.getNamedProp('80DA', ps.PSETID_ADDRESS)

    @functools.cached_property
    def homeAddressLocality(self) -> Optional[str]:
        """
        The locality or city portion of the contact's home address.
        """
        return self.getStringStream('__substg1.0_3A59')

    @functools.cached_property
    def homeAddressPostalCode(self) -> Optional[str]:
        """
        The postal code portion of the contact's home address.
        """
        return self.getStringStream('__substg1.0_3A5B')

    @functools.cached_property
    def homeAddressPostOfficeBox(self) -> Optional[str]:
        """
        The number or identifier of the contact's home post office box.
        """
        return self.getStringStream('__substg1.0_3A5E')

    @functools.cached_property
    def homeAddressStateOrProvince(self) -> Optional[str]:
        """
        The state or province portion of the contact's home address.
        """
        return self.getStringStream('__substg1.0_3A5C')

    @functools.cached_property
    def homeAddressStreet(self) -> Optional[str]:
        """
        The street portion of the contact's home address.
        """
        return self.getStringStream('__substg1.0_3A5D')

    @functools.cached_property
    def homeFax(self) -> Optional[_FAX_DICT]:
        """
        Returns a ``dict`` of the data for the home fax.

        Returns ``None`` if no fields are set.

        Keys are "address_type", "email_address", "number",
        "original_display_name", and "original_entry_id".
        """
        data = {
            'address_type': self.homeFaxAddressType,
            'email_address': self.homeFaxEmailAddress,
            'number': self.homeFaxNumber,
            'original_display_name': self.homeFaxOriginalDisplayName,
            'original_entry_id': self.homeFaxOriginalEntryId,
        }
        return data if any(data[x] for x in data) else None

    @functools.cached_property
    def homeFaxAddressType(self) -> Optional[str]:
        """
        The type of address for the fax. MUST be set to "FAX" if present.
        """
        return self.getNamedProp('80D2', ps.PSETID_ADDRESS)

    @functools.cached_property
    def homeFaxEmailAddress(self) -> Optional[str]:
        """
        Contains a user-readable display name, followed by the "@" character,
        followed by a fax number.
        """
        return self.getNamedProp('80D3', ps.PSETID_ADDRESS)

    @functools.cached_property
    def homeFaxNumber(self) -> Optional[str]:
        """
        Contains the number of the contact's home fax.
        """
        return self.getStringStream('__substg1.0_3A25')

    @functools.cached_property
    def homeFaxOriginalDisplayName(self) -> Optional[str]:
        """
        The normalized subject for the contact.
        """
        return self.getNamedProp('80D4', ps.PSETID_ADDRESS)

    @functools.cached_property
    def homeFaxOriginalEntryId(self) -> Optional[EntryID]:
        """
        The one-off EntryID corresponding to this fax address.
        """
        return self.getNamedAs('80D5', ps.PSETID_ADDRESS, EntryID.autoCreate)

    @functools.cached_property
    def homeTelephoneNumber(self) -> Optional[str]:
        """
        The number of the contact's home telephone.
        """
        return self.getStringStream('__substg1.0_3A09')

    @functools.cached_property
    def homeTelephone2Number(self) -> Optional[Union[str, List[str]]]:
        """
        The number(s) of the contact's second home telephone.
        """
        return self.getSingleOrMultipleString('__substg1.0_3A2F')

    @functools.cached_property
    def initials(self) -> Optional[str]:
        """
        The initials of the contact.
        """
        return self.getStringStream('__substg1.0_3A0A')

    @functools.cached_property
    def instantMessagingAddress(self) -> Optional[str]:
        """
        The instant messaging address of the contact.
        """
        return self.getNamedProp('8062', ps.PSETID_ADDRESS)

    @functools.cached_property
    def isContactLinked(self) -> bool:
        """
        Whether the contact is linked to other contacts.
        """
        return bool(self.getNamedProp('80E0', ps.PSETID_ADDRESS))

    @functools.cached_property
    def isdnNumber(self) -> Optional[str]:
        """
        The Integrated Services Digital Network (ISDN) telephone number of the
        contact.
        """
        return self.getStringStream('__substg1.0_3A2D')

    @functools.cached_property
    def jobTitle(self) -> Optional[str]:
        """
        The job title of the contact.
        """
        return self.getStringStream('__substg1.0_3A17')

    @functools.cached_property
    def language(self) -> Optional[str]:
        """
        The language that the contact uses.
        """
        return self.getStringStream('__substg1.0_3A0C')

    @functools.cached_property
    def lastModifiedBy(self) -> Optional[str]:
        """
        The name of the last user to modify the contact file.
        """
        return self.getStringStream('__substg1.0_3FFA')

    @functools.cached_property
    def location(self) -> Optional[str]:
        """
        The location of the contact.

        For example, this could be the building or office number of the contact.
        """
        return self.getStringStream('__substg1.0_3A0D')

    @functools.cached_property
    def mailAddress(self) -> Optional[str]:
        """
        The complete mail address of the contact.
        """
        return self.getStringStream('__substg1.0_3A15')

    @functools.cached_property
    def mailAddressCountry(self) -> Optional[str]:
        """
        The country portion of the contact's mail address.
        """
        return self.getStringStream('__substg1.0_3A26')

    @functools.cached_property
    def mailAddressCountryCode(self) -> Optional[str]:
        """
        The country code portion of the contact's mail address.
        """
        return self.getNamedProp('80DD', ps.PSETID_ADDRESS)

    @functools.cached_property
    def mailAddressLocality(self) -> Optional[str]:
        """
        The locality or city portion of the contact's mail address.
        """
        return self.getStringStream('__substg1.0_3A27')

    @functools.cached_property
    def mailAddressPostalCode(self) -> Optional[str]:
        """
        The postal code portion of the contact's mail address.
        """
        return self.getStringStream('__substg1.0_3A2A')

    @functools.cached_property
    def mailAddressPostOfficeBox(self) -> Optional[str]:
        """
        The number or identifier of the contact's mail post office box.
        """
        return self.getStringStream('__substg1.0_3A2B')

    @functools.cached_property
    def mailAddressStateOrProvince(self) -> Optional[str]:
        """
        The state or province portion of the contact's mail address.
        """
        return self.getStringStream('__substg1.0_3A28')

    @functools.cached_property
    def mailAddressStreet(self) -> Optional[str]:
        """
        The street portion of the contact's mail address.
        """
        return self.getStringStream('__substg1.0_3A29')

    @functools.cached_property
    def managerName(self) -> Optional[str]:
        """
        The name of the contact's manager.
        """
        return self.getStringStream('__substg1.0_3A4E')

    @functools.cached_property
    def middleName(self) -> Optional[str]:
        """
        The middle name(s) of the contact.
        """
        return self.getStringStream('__substg1.0_3A44')

    @functools.cached_property
    def mobileTelephoneNumber(self) -> Optional[str]:
        """
        The mobile telephone number of the contact.
        """
        return self.getStringStream('__substg1.0_3A1C')

    @functools.cached_property
    def nickname(self) -> Optional[str]:
        """
        The nickname of the contanct.
        """
        return self.getStringStream('__substg1.0_3A4F')

    @functools.cached_property
    def officeLocation(self) -> Optional[str]:
        """
        The location of the office that the contact works in.
        """
        return self.getStringStream('__substg1.0_3A19')

    @functools.cached_property
    def organizationalIDNumber(self) -> Optional[str]:
        """
        The organizational ID number for the contact, such as an employee ID
        number.
        """
        return self.getStringStream('__substg1.0_3A10')

    @functools.cached_property
    def oscSyncEnabled(self) -> bool:
        """
        Whether contact synchronization with an external source (such as a
        social networking site) is handled by the server.
        """
        return bool(self.getPropertyVal('7C24000B'))

    @functools.cached_property
    def otherAddress(self) -> Optional[str]:
        """
        The complete other address of the contact.
        """
        return self.getNamedProp('801C', ps.PSETID_ADDRESS)

    @functools.cached_property
    def otherAddressCountry(self) -> Optional[str]:
        """
        The country portion of the contact's other address.
        """
        return self.getStringStream('__substg1.0_3A60')

    @functools.cached_property
    def otherAddressCountryCode(self) -> Optional[str]:
        """
        The country code portion of the contact's other address.
        """
        return self.getNamedProp('80DC', ps.PSETID_ADDRESS)

    @functools.cached_property
    def otherAddressLocality(self) -> Optional[str]:
        """
        The locality or city portion of the contact's other address.
        """
        return self.getStringStream('__substg1.0_3A5F')

    @functools.cached_property
    def otherAddressPostalCode(self) -> Optional[str]:
        """
        The postal code portion of the contact's other address.
        """
        return self.getStringStream('__substg1.0_3A61')

    @functools.cached_property
    def otherAddressPostOfficeBox(self) -> Optional[str]:
        """
        The number or identifier of the contact's other post office box.
        """
        return self.getStringStream('__substg1.0_3A64')

    @functools.cached_property
    def otherAddressStateOrProvince(self) -> Optional[str]:
        """
        The state or province portion of the contact's other address.
        """
        return self.getStringStream('__substg1.0_3A62')

    @functools.cached_property
    def otherAddressStreet(self) -> Optional[str]:
        """
        The street portion of the contact's other address.
        """
        return self.getStringStream('__substg1.0_3A63')

    @functools.cached_property
    def otherTelephoneNumber(self) -> Optional[str]:
        """
        Contains the number of the contact's other telephone.
        """
        return self.getStringStream('__substg1.0_3A1F')

    @functools.cached_property
    def pagerTelephoneNumber(self) -> Optional[str]:
        """
        The contact's pager telephone number.
        """
        return self.getStringStream('__substg1.0_3A21')

    @functools.cached_property
    def personalHomePage(self) -> Optional[str]:
        """
        The contact's personal web page UL.
        """
        return self.getStringStream('__substg1.0_3A50')

    @functools.cached_property
    def phoneticCompanyName(self) -> Optional[str]:
        """
        The phonetic pronunciation of the contact's company name.
        """
        return self.getNamedProp('802E', ps.PSETID_ADDRESS)

    @functools.cached_property
    def phoneticGivenName(self) -> Optional[str]:
        """
        The phonetic pronunciation of the contact's given name.
        """
        return self.getNamedProp('802C', ps.PSETID_ADDRESS)

    @functools.cached_property
    def phoneticSurname(self) -> Optional[str]:
        """
        The phonetic pronunciation of the given name of the contact.
        """
        return self.getNamedProp('802D', ps.PSETID_ADDRESS)

    @functools.cached_property
    def postalAddressID(self) -> PostalAddressID:
        """
        Indicates which physical address is the Mailing Address for this
        contact.
        """
        return PostalAddressID(self.getNamedProp('8022', ps.PSETID_ADDRESS, 0))

    @functools.cached_property
    def primaryFax(self) -> Optional[_FAX_DICT]:
        """
        Returns a ``dict`` of the data for the primary fax.

        Returns ``None`` if no fields are set.

        Keys are "address_type", "email_address", "number",
        "original_display_name", and "original_entry_id".
        """
        data = {
            'address_type': self.primaryFaxAddressType,
            'email_address': self.primaryFaxEmailAddress,
            'number': self.primaryFaxNumber,
            'original_display_name': self.primaryFaxOriginalDisplayName,
            'original_entry_id': self.primaryFaxOriginalEntryId,
        }
        return data if any(data[x] for x in data) else None

    @functools.cached_property
    def primaryFaxAddressType(self) -> Optional[str]:
        """
        The type of address for the fax. MUST be set to "FAX" if present.
        """
        return self.getNamedProp('80B2', ps.PSETID_ADDRESS)

    @functools.cached_property
    def primaryFaxEmailAddress(self) -> Optional[str]:
        """
        Contains a user-readable display name, followed by the "@" character,
        followed by a fax number.
        """
        return self.getNamedProp('80B3', ps.PSETID_ADDRESS)

    @functools.cached_property
    def primaryFaxNumber(self) -> Optional[str]:
        """
        Contains the number of the contact's primary fax.
        """
        return self.getStringStream('__substg1.0_3A23')

    @functools.cached_property
    def primaryFaxOriginalDisplayName(self) -> Optional[str]:
        """
        The normalized subject for the contact.
        """
        return self.getNamedProp('80B4', ps.PSETID_ADDRESS)

    @functools.cached_property
    def primaryFaxOriginalEntryId(self) -> Optional[EntryID]:
        """
        The one-off EntryID corresponding to this fax address.
        """
        return self.getNamedAs('80B5', ps.PSETID_ADDRESS, EntryID.autoCreate)

    @functools.cached_property
    def primaryTelephoneNumber(self) -> Optional[str]:
        """
        Contains the number of the contact's primary telephone.
        """
        return self.getStringStream('__substg1.0_3A1A')

    @functools.cached_property
    def profession(self) -> Optional[str]:
        """
        The profession of the contact.
        """
        return self.getStringStream('__substg1.0_3A46')

    @functools.cached_property
    def radioTelephoneNumber(self) -> Optional[str]:
        """
        Contains the number of the contact's radio telephone.
        """
        return self.getStringStream('__substg1.0_3A1D')

    @functools.cached_property
    def referenceEntryID(self) -> Optional[EntryID]:
        """
        Contains a value that is equal to the value of the EntryID of the
        Contact object unless the Contact object is a copy of an earlier
        original.
        """
        return self.getNamedAs('85BD', ps.PSETID_COMMON, EntryID.autoCreate)

    @functools.cached_property
    def referredByName(self) -> Optional[str]:
        """
        The name of the person who referred this contact to the user.
        """
        return self.getStringStream('__substg1.0_3A47')

    @functools.cached_property
    def spouseName(self) -> Optional[str]:
        """
        The name of the contact's spouse.
        """
        return self.getStringStream('__substg1.0_3A48')

    @functools.cached_property
    def surname(self) -> Optional[str]:
        """
        The surname of the contact.
        """
        return self.getStringStream('__substg1.0_3A11')

    @functools.cached_property
    def tddTelephoneNumber(self) -> Optional[str]:
        """
        The telephone number for the contact's text telephone (TTY) or
        telecommunication device for the deaf (TDD).
        """
        return self.getStringStream('__substg1.0_3A4B')

    @functools.cached_property
    def telexNumber(self) -> Optional[Union[str, List[str]]]:
        """
        The contact's telex number(s).
        """
        return self.getSingleOrMultipleString('__substg1.0_3A2C')

    @functools.cached_property
    def userX509Certificate(self) -> Optional[List[bytes]]:
        """
        A list of certificates for the contact.
        """
        return self.getMultipleBinary('3A70')

    @functools.cached_property
    def weddingAnniversary(self) -> Optional[datetime.datetime]:
        """
        The wedding anniversary of the contact at 11:59 UTC.
        """
        return self.getPropertyVal('3A410040')

    @functools.cached_property
    def weddingAnniversaryEventEntryID(self) -> Optional[EntryID]:
        """
        The EntryID of an optional Appointement object that represents the
        contact's wedding anniversary.
        """
        return self.getNamedAs('804E', ps.PSETID_ADDRESS, EntryID.autoCreate)

    @functools.cached_property
    def weddingAnniversaryLocal(self) -> Optional[datetime.datetime]:
        """
        The wedding anniversary of the contact at 0:00 in the client's local
        time zone.
        """
        return self.getNamedProp('80DF', ps.PSETID_ADDRESS)

    @functools.cached_property
    def webpageUrl(self) -> Optional[str]:
        """
        The contact's business web page url. SHOULD be the same as businessUrl.
        """
        return self.getNamedProp('802B', ps.PSETID_ADDRESS)

    @functools.cached_property
    def workAddress(self) -> Optional[str]:
        """
        The complete work address of the contact.
        """
        return self.getNamedProp('801B', ps.PSETID_ADDRESS)

    @functools.cached_property
    def workAddressCountry(self) -> Optional[str]:
        """
        The country portion of the contact's work address.
        """
        return self.getNamedProp('8049', ps.PSETID_ADDRESS)

    @functools.cached_property
    def workAddressCountryCode(self) -> Optional[str]:
        """
        The country code portion of the contact's work address.
        """
        return self.getNamedProp('80DB', ps.PSETID_ADDRESS)

    @functools.cached_property
    def workAddressLocality(self) -> Optional[str]:
        """
        The locality or city portion of the contact's work address.
        """
        return self.getNamedProp('8046', ps.PSETID_ADDRESS)

    @functools.cached_property
    def workAddressPostalCode(self) -> Optional[str]:
        """
        The postal code portion of the contact's work address.
        """
        return self.getNamedProp('8048', ps.PSETID_ADDRESS)

    @functools.cached_property
    def workAddressPostOfficeBox(self) -> Optional[str]:
        """
        The number or identifier of the contact's work post office box.
        """
        return self.getNamedProp('804A', ps.PSETID_ADDRESS)

    @functools.cached_property
    def workAddressStateOrProvince(self) -> Optional[str]:
        """
        The state or province portion of the contact's work address.
        """
        return self.getNamedProp('8047', ps.PSETID_ADDRESS)

    @functools.cached_property
    def workAddressStreet(self) -> Optional[str]:
        """
        The street portion of the contact's work address.
        """
        return self.getNamedProp('8045', ps.PSETID_ADDRESS)
