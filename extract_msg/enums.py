import enum

from typing import Set


class AddressBookType(enum.Enum):
    """
    The type of object that an address book entry ID represents. MUST be one of
    these or it is invalid.
    """
    LOCAL_MAIL_USER = 0x000
    DISTRIBUTION_LIST = 0x001
    BULLETIN_BOARD_OR_PUBLIC_FOLDER = 0x002
    AUTOMATED_MAILBOX = 0x003
    ORGANIZATIONAL_MAILBOX = 0x004
    PRIVATE_DISTRIBUTION_LIST = 0x005
    REMOTE_MAIL_USER = 0x006
    CONTAINER = 0x100
    TEMPLATE = 0x101
    ONE_OFF_USER = 0x102
    SEARCH = 0x200

class AttachErrorBehavior(enum.Enum):
    """
    The behavior to follow when handling an error in an attachment.
    THROW: Throw the exception regardless of type.
    NOT_IMPLEMENTED: Silence the exception for NotImplementedError.
    BROKEN: Silence the exception for NotImplementedError and for broken
        attachments.
    """
    THROW = 0
    NOT_IMPLEMENTED = 1
    BROKEN = 2

class AttachmentType(enum.Enum):
    """
    The type represented by the attachment.
    """
    DATA = 0
    MSG = 1
    WEB = 2
    SIGNED = 3

class DeencapType(enum.Enum):
    """
    Enum to specify to custom deencapsulation functions the type of data being
    requested.
    """
    PLAIN = 0
    HTML = 1

class DisplayType(enum.Enum):
    MAILUSER = 0x0000
    DISTLIST = 0x0001
    FORUM = 0x0002
    AGENT = 0x0003
    ORGANIZATION = 0x0004
    PRIVATE_DISTLIST = 0x0005
    REMOTE_MAILUSER = 0x0006
    CONTAINER = 0x0100
    TEMPLATE = 0x0101
    ADDRESS_TEMPLATE = 0x0102
    SEARCH = 0x0200

class ElectronicAddressProperties(enum.Enum):
    @classmethod
    def fromBits(cls, value : int) -> Set['ElectronicAddressProperties']:
        """
        Converts an int, with the left most bit referring to 0x00000000, to a
        set of this enum.

        :raises ValueError: The value was less than 0.
        """
        if value < 0:
            raise ValueError('Value must not be negative.')
        # This is a quick compressed way to convert the bits in the int into
        # a tuple of instances of this class should any bit be a 1.
        return tuple(cls(int(x)) for index, val in enumerate(bin(value)[:1:-1]) if val == '1')

    EMAIL_1 = 0x00000000
    EMAIL_2 = 0x00000001
    EMAIL_3 = 0x00000002
    BUSINESS_FAX = 0x00000003
    HOME_FAX = 0x00000004
    PRIMARY_FAX = 0x00000005

class EntryIDType(enum.Enum):
    """
    Converts a UID to the type of Entry ID structure.
    """
    def toHex(self):
        """
        Converts an EntryIDType to it's hex equivelent.
        """
        return EntryIDTypeHex[self.name]

    PUBLIC_MESSAGE_STORE = b'\x1A\x44\x73\x90\xAA\x66\x11\xCD\x9B\xC8\x00\xAA\x00\x2F\xC4\x5A'
    ADDRESS_BOOK_RECIPIENT = b'\xDC\xA7\x40\xC8\xC0\x42\x10\x1A\xB4\xB9\x08\x00\x2B\x2F\xE1\x82'
    ONE_OFF_RECIPIENT = b'\x81\x2B\x1F\xA4\xBE\xA3\x10\x19\x9D\x6E\x00\xDD\x01\x0F\x54\x02'
    # Contact address or personal distribution list recipient.
    CA_OR_PDL_RECIPIENT = b'\xFE\x42\xAA\x0A\x18\xC7\x1A\x10\xE8\x85\x0B\x65\x1C\x24\x00\x00'
    NNTP_NEWSGROUP_FOLDER = b'\x38\xA1\xBB\x10\x05\xE5\x10\x1A\xA1\xBB\x08\x00\x2B\x2A\x56\xC2'
    # [MS-OXOCNTC] WrappedEntryId Structure.
    WRAPPED = b'\xC0\x91\xAD\xD3\x51\x9D\xCF\x11\xA4\xA9\x00\xAA\x00\x47\xFA\xA4'

class EntryIDTypeHex(enum.Enum):
    """
    Converts a UID to the type of Entry ID structure. Uses a hex string instead
    of bytes for the value.
    """
    def toRaw(self):
        """
        Converts and EntryIDTypeHex to it's raw equivelent.
        """
        return EntryIDType[self.name]

    PUBLIC_MESSAGE_STORE = '1A447390AA6611CD9BC800AA002FC45A'
    ADDRESS_BOOK_RECIPIENT = 'DCA740C8C042101AB4B908002B2FE182'
    ONE_OFF_RECIPIENT = '812B1FA4BEA310199D6E00DD010F5402'
    # Contact address or personal distribution list recipient.
    CA_OR_PDL_RECIPIENT = 'FE42AA0A18C71A10E8850B651C240000'
    NNTP_NEWSGROUP_FOLDER = '38A1BB1005E5101AA1BB08002B2A56C2'
    # [MS-OXOCNTC] WrappedEntryId Structure.
    WRAPPED = 'C091ADD3519DCF11A4A900AA0047FAA4'

class Gender(enum.Enum):
    # Seems rather binary, which is less than ideal. We are directly using the
    # terms used by the documentation.
    UNSPECIFIED = 0x0000
    FEMALE = 0x0001
    MALE = 0x0002

class Guid(enum.Enum):
    PS_MAPI = '{00020328-0000-0000-C000-000000000046}'
    PS_PUBLIC_STRINGS = '{00020329-0000-0000-C000-000000000046}'
    PSETID_COMMON = '{00062008-0000-0000-C000-000000000046}'
    PSETID_ADDRESS = '{00062004-0000-0000-C000-000000000046}'
    PS_INTERNET_HEADERS = '{00020386-0000-0000-C000-000000000046}'
    PSETID_APPOINTMENT = '{00062002-0000-0000-C000-000000000046}'
    PSETID_MEETING = '{6ED8DA90-450B-101B-98DA-00AA003F1305}'
    PSETID_LOG = '{0006200A-0000-0000-C000-000000000046}'
    PSETID_MESSAGING = '{41F28F13-83F4-4114-A584-EEDB5A6B0BFF}'
    PSETID_NOTE = '{0006200E-0000-0000-C000-000000000046}'
    PSETID_POSTRSS = '{00062041-0000-0000-C000-000000000046}'
    PSETID_TASK = '{00062003-0000-0000-C000-000000000046}'
    PSETID_UNIFIEDMESSAGING = '{4442858E-A9E3-4E80-B900-317A210CC15B}'
    PSETID_AIRSYNC = '{71035549-0739-4DCB-9163-00F0580DBBDF}'
    PSETID_SHARING = '{00062040-0000-0000-C000-000000000046}'
    PSETID_XMLEXTRACTEDENTITIES = '{23239608-685D-4732-9C55-4C95CB4E8E33}'
    PSETID_ATTACHMENT = '{96357F7F-59E1-47D0-99A7-46515C183B54}'

class Importance(enum.Enum):
    LOW = 0
    MEDIUM = 1
    IMPORTANCE_HIGH = 2

class Intelligence(enum.Enum):
    DUMB = 0
    SMART = 1

class MacintoshEncoding(enum.Enum):
    """
    The encoding to use for Macintosh-specific data attachments.
    """
    BIN_HEX = 0
    UUENCODE = 1
    APPLE_SINGLE = 2
    APPLE_DOUBLE = 3

class MessageFormat(enum.Enum):
    TNEF = 0
    MIME = 1

class NamedPropertyType(enum.Enum):
    NUMERICAL_NAMED = 0
    STRING_NAMED = 1

class OORBodyFormat(enum.Enum):
    """
    The body format for One Off Recipients.
    """
    TEXT_ONLY = 0b0011
    HTML_ONLY = 0b0111
    TEXT_AND_HTML = 0b1011
    # This one isn't actually listed in the documentation, but I've seen it and
    # this is my best guess for what a format of `0` is meant to mean. This will
    # also prevent the code from failing on a 0 format.
    UNSPECIFIED = 0b0000

class PostalAddressID(enum.Enum):
    UNSPECIFIED = 0x00000000
    HOME = 0x00000001
    WORK = 0x00000002
    OTHER = 0x00000003

class Priority(enum.Enum):
    URGENT = 0x00000001
    NORMAL = 0x00000000
    NOT_URGENT = 0xFFFFFFFF

class PropertiesType(enum.Enum):
    """
    The type of the properties instance.
    """
    MESSAGE = 0
    MESSAGE_EMBED = 1
    ATTACHMENT = 2
    RECIPIENT = 3

class RecipientRowFlagType(enum.Enum):
    NOTYPE = 0x0
    X500DN = 0x1
    MSMAIL = 0x2
    SMTP = 0x3
    FAX = 0x4
    PROFESSIONALOFFICESYSTEM = 0x5
    PERSONALDESTRIBUTIONLIST1 = 0x6
    PERSONALDESTRIBUTIONLIST2 = 0x7

class RecipientType(enum.Enum):
    """
    The type of recipient.
    """
    SENDER = 0
    TO = 1
    CC = 2
    BCC = 3

class RuleActionType(enum.Enum):
    OP_MOVE = 0x01
    OP_COPY = 0x02
    OP_REPLY = 0x03
    OP_OOF_REPLY = 0x04
    OP_DEFER_ACTION = 0x05
    OP_BOUNCE = 0x06
    OP_FORWARD = 0x07
    OP_DELEGATE = 0x08
    OP_TAG = 0x09
    OP_DELETE = 0x0A
    OP_MARK_AS_READ = 0x0B

class Sensitivity(enum.Enum):
    NORMAL = 0
    PERSONAL = 1
    PRIVATE = 2
    CONFIDENTIAL = 3

class TaskAcceptance(enum.Enum):
    """
    The acceptance state of the task.
    """
    NOT_ASSIGNED = 0x00000000
    UNKNOWN = 0x00000001
    ACCEPTED = 0x00000002
    REJECTED = 0x00000003

class TaskHistory(enum.Enum):
    """
    The type of the last change to the Task object.
    """
    NONE = 0x00000000
    ACCEPTED = 0x00000001
    REJECTED = 0x00000002
    OTHER = 0x00000003
    DUE_DATE_CHANGED = 0x00000004
    ASSIGNED = 0x00000005

class TaskMode(enum.Enum):
    """
    The mode of the Task object used in task communication (PidLidTaskMode).

    UNASSIGNED: The Task object is not assigned.
    EMBEDDED_REQUEST: The Task object is embedded in a task request.
    ACCEPTED: The Task object has been accepted by the task assignee.
    REJECTED: The Task object was rejected by the task assignee.
    EMBEDDED_UPDATE: The Task object is embedded in a task update.
    SELF_ASSIGNED: The Task object was assigned to the task assigner
        (self-delegation).
    """
    UNASSIGNED = 0
    EMBEDDED_REQUEST = 1
    ACCEPTED = 2
    REJECTED = 3
    EMBEDDED_UPDATE = 4
    SELF_ASSIGNED = 5

class TaskOwnership(enum.Enum):
    """
    The role of the current user relative to the Task object.

    NOT_ASSIGNED: The Task object is not assigned.
    ASSIGNERS_COPY: The Task object is the task assigner's copy of the Task
        object.
    ASSIGNEES_COPY: The Task object is the task assignee's copy of the Task
        object.
    """
    NOT_ASSIGNED = 0x00000000
    ASSIGNERS_COPY = 0x00000001
    ASSIGNEES_COPY = 0x00000002

class TaskStatus(enum.Enum):
    """
    The status of a task object (PidLidTaskStatus).

    NOT_STARTED: The user has not started the task.
    IN_PROGRESS: The users's work on the Task object is in progress.
    COMPLETE: The user's work on the Task object is complete.
    WAITING_ON_OTHER: The user is waiting on somebody else.
    DEFERRED: The user has deffered work on the Task object.
    """
    NOT_STARTED = 0x00000000
    IN_PROGRESS = 0x00000001
    COMPLETE = 0x00000002
    WAITING_ON_OTHER = 0x00000003
    DEFERRED = 0x00000004
