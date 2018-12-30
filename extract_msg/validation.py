import copy

import olefile

from extract_msg.message import Message
from extract_msg.utils import get_full_name, has_len


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

validation_dict_msg = {
    '001F/001E': None,
    'header': {
        'exists': False,
        'not empty': False,
    },
    'body': {
        'exists': False,
        'not empty': False,
    },
    'attachments': None,
    'recipients': None,
},

validation_dict_attachment = {

}

validation_dict_recipient = {

}


def validate_msg(instance):

def validate_attachment(instance):

def validate_recipient(instance):

def validate(msg):
    validation_dict = {
        'input': {
            'class': get_full_name(msg), # Get the full name of the class
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
        except:
            pass;
        else:
            validation_dict['msg']['initializes'] = True
            for x in msg_list:
                #TODO this needs to be done is some way that makes it so that the msg structure can easily be calculated and the dicts can be properly embedded without using recursive functions
    return validation_dict
