import copy

import olefile

from extract_msg.message import Message
from extract_msg.utils import get_full_name, has_len

'': {
    'exists': False,
    'not empty': False,
},


validation_dict_base = {
    'input': {
        'class': 'None',
        'has_len': False,
        'len': 0,
    },
    'olefile': {
        'valid': False,
    },
    'msg': None,
}

validation_dict_attachment = {
    'long filename': {
        'exists': False,
        'not empty': False,
    },
    'short filename': {
        'exists': False,
        'not empty': False,
    },
    'content id': {
        'exists': False,
        'not empty': False,
    },
    '': {
        'exists': False,
        'not empty': False,
    },
    'type': None,

}

validation_dict_recipient = {

}
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

def validate_msg(instance):
    return {
        '001F/001E': None,
        'header': get_string_details(instance, '__substg1.0_007D'),
        'body': get_string_details('__substg1.0_1000'),
        'html body': get_stream_details('__substg1.0_10130102'),
        'rtf body': get_stream_details('__substg1.0_10090102'),
        'date': instance.date,
        'attachments': {x: validate_attachment(y) for x, y in enumerate instance.attachments},
        'recipients': {x: validate_recipient(y) for x, y in enumerate instance.recipients},
    }

def validate_attachment(instance):

def validate_recipient(instance):
    return {
        '': '',
        'stream 3003': get_email_details(),
    }

def validate(msg):
    validation_dict = {
        'input': {
            'class': get_full_class_name(msg), # Get the full name of the class
            'has_len': has_len(msg), # Does the input have a __len__ attribute?
            'len': len(msg) if has_len(msg) else None, # If input has __len__, put the value here
        },
        'olefile': {
            'valid': olefile.isolefile(msg),
        },
    }
    if validation_dict['olefile']['valid']:
        validation_dict['msg'] = {
            'initializes': False,
        }
        try:
            msg_list = [Message(msg)]
            msg_dicts = []
        except NotImplementedError:
            # Should we have a special procedure for handling it if we get "not implemented"?
            pass
        except:
            pass
        else:
            validation_dict['msg']['initializes'] = True
            for x in msg_list:
                #TODO this needs to be done is some way that makes it so that the msg structure can easily be calculated and the dicts can be properly embedded without using recursive functions
    return validation_dict
