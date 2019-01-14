import copy

import olefile

from extract_msg.message import Message
from extract_msg.utils import get_full_class_name, has_len


def get_string_details(instance, stream):
    return {
        'exists': instance.sExists(stream),
        'not empty': False if not instance.sExists(stream) else len(instance._getStringStream(stream)) > 0,
    }


def get_stream_details(instance, stream):
    return {
        'exists': instance.Exists(stream),
        'not empty': False if not instance.Exists(stream) else len(instance._getStream(stream)) > 0,
    }


def get_email_details(instance, stream):
    return {
        'exists': instance.sExists(stream),
        'not empty': False if not instance.sExists(stream) else len(instance._getStringStream(stream)) > 0,
        'valid email address': False if not instance.sExists(stream) else u'@' in instance._getStringStream(stream),
    }


def string_FE(instance):
    temp = '001E'
    if instance.mainProperties.has_key('340D0003'):
        temp = '001F' if instance.mainProperties['340D0003'].value & 0x40000 else '001E'
    tempnot = '001F' if temp == '001E' else '001E'
    confirmation = all(not x[-1].endswith(tempnot) for x in instance.listDir())
    if confirmation:
        temp += ', but ' + tempnot + ' was detected.'
    return temp


def validate_msg(instance):
    return {
        '001F/001E': string_FE(instance),
        'header': get_string_details(instance, '__substg1.0_007D'),
        'body': get_string_details(instance, '__substg1.0_1000'),
        'html body': get_stream_details(instance, '__substg1.0_10130102'),
        'rtf body': get_stream_details(instance, '__substg1.0_10090102'),
        'date': instance.date,
        'attachments': {x: validate_attachment(y) for x, y in enumerate(instance.attachments)},
        'recipients': {x: validate_recipient(y) for x, y in enumerate(instance.recipients)},
    }


def validate_attachment(instance):
    temp = {
        'long filename': get_string_details(instance, '__substg1.0_3707'),
        'short filename': get_string_details(instance, '__substg1.0_3704'),
        'content id': get_string_details(instance, '__substg1.0_3712'),
        'type': instance.type,
    }
    if temp['type'] == 'msg':
        temp['msg'] = validate_msg(instance.data)
    return temp


def validate_recipient(instance):
    return {
        'type': instance.type,
        'stream 3003': get_email_details(instance, '__substg1.0_3003'),
        'stream 39FE': get_email_details(instance, '__substg1.0_39FE'),
    }


def validate(msg):
    validation_dict = {
        'input': {
            'class': get_full_class_name(msg), # Get the full name of the class
            'has_len': has_len(msg), # Does the input have a __len__ attribute?
            'len': len(msg) if has_len(msg) else None, # If input has __len__, put the value here
        },
        'olefile': {
            'valid': olefile.isOleFile(msg),
        },
    }
    if validation_dict['olefile']['valid']:
        validation_dict['message'] = {
            'initializes': False,
        }
        try:
            msg_instance = Message(msg)
        except NotImplementedError:
            # Should we have a special procedure for handling it if we get "not implemented"?
            pass
        except:
            pass
        else:
            validation_dict['message']['initializes'] = True
            validation_dict['message']['msg'] = validate_msg(msg_instance)
    return validation_dict
