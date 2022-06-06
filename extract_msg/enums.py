import enum


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

class NamedPropertyType(enum.Enum):
    NUMERICAL_NAMED = 0
    STRING_NAMED = 1

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

class RecipientType(enum.Enum):
    """
    The type of recipient.
    """
    SENDER = 0
    TO = 1
    CC = 2
    BCC = 3

class Sensitivity(enum.Enum):
    NORMAL = 0
    PERSONAL = 1
    PRIVATE = 2
    CONFIDENTIAL = 3
