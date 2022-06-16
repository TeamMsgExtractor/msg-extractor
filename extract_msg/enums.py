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
