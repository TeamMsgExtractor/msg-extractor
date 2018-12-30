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


def validate_msg(instance):
    return {
        '001F/001E': None,
        'header': {
            'exists': False,
            'not empty': False,
        },
        'body': {
            'exists': instance.sExists('__substg1.0_')
            'not empty': False,
        },
        'date': instance.date,
        'attachments': {x: validate_attachment(y) for x, y in enumerate instance.attachments},
        'recipients': {x: validate_recipient(y) for x, y in enumerate instance.recipients},
    }

def validate_attachment(instance):

def validate_recipient(instance):
    return {
        '': '',
        'stream 3003': {
            'exists': instance.msg.sExists([instance.dir, '__substg1.0_3003']),
            'not empty': False,
            'valid email address':
        },
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
