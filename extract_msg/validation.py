import olefile

from .message import Message
from .utils import getFullClassName, hasLen


def getEmailDetails(instance, stream):
    return {
        'exists': instance.sExists(stream),
        'not empty': False if not instance.sExists(stream) else len(instance._getStringStream(stream)) > 0,
        'valid email address': False if not instance.sExists(stream) else u'@' in instance._getStringStream(stream),
    }

def getStreamDetails(instance, stream):
    return {
        'exists': instance.exists(stream),
        'not empty': False if not instance.exists(stream) else len(instance._getStream(stream)) > 0,
    }

def getStringDetails(instance, stream):
    return {
        'exists': instance.sExists(stream),
        'not empty': False if not instance.sExists(stream) else len(instance._getStringStream(stream)) > 0,
    }

def stringFE(instance):
    temp = '001E'
    if instance.mainProperties.has_key('340D0003'):
        temp = '001F' if instance.mainProperties['340D0003'].value & 0x40000 else '001E'
    tempnot = '001F' if temp == '001E' else '001E'
    confirmation = all(not x[-1].endswith(tempnot) for x in instance.listDir())
    if confirmation:
        temp += ', but ' + tempnot + ' was detected.'
    return temp


def validate(msg):
    validationDict = {
        'input': {
            'class': getFullClassName(msg), # Get the full name of the class
            'hasLen': hasLen(msg), # Does the input have a __len__ attribute?
            'len': len(msg) if hasLen(msg) else None, # If input has __len__, put the value here
        },
        'olefile': {
            'valid': olefile.isOleFile(msg),
        },
    }
    if validationDict['olefile']['valid']:
        validationDict['message'] = {
            'initializes': False,
        }
        try:
            msgInstance = Message(msg)
        except NotImplementedError:
            # Should we have a special procedure for handling it if we get "not implemented"?
            pass
        except:
            pass
        else:
            validationDict['message']['initializes'] = True
            validationDict['message']['msg'] = validateMsg(msgInstance)
    return validationDict

def validateAttachment(instance):
    temp = {
        'long filename': getStringDetails(instance, '__substg1.0_3707'),
        'short filename': getStringDetails(instance, '__substg1.0_3704'),
        'content id': getStringDetails(instance, '__substg1.0_3712'),
        'type': instance.type,
    }
    if temp['type'] == 'msg':
        temp['msg'] = validateMsg(instance.data)
    return temp

def validateMsg(instance):
    return {
        '001F/001E': stringFE(instance),
        'header': getStringDetails(instance, '__substg1.0_007D'),
        'body': getStringDetails(instance, '__substg1.0_1000'),
        'html body': getStreamDetails(instance, '__substg1.0_10130102'),
        'rtf body': getStreamDetails(instance, '__substg1.0_10090102'),
        'date': instance.date,
        'attachments': {x: validateAttachment(y) for x, y in enumerate(instance.attachments)},
        'recipients': {x: validateRecipient(y) for x, y in enumerate(instance.recipients)},
    }

def validateRecipient(instance):
    return {
        'type': instance.type,
        'stream 3003': getEmailDetails(instance, '__substg1.0_3003'),
        'stream 39FE': getEmailDetails(instance, '__substg1.0_39FE'),
    }
