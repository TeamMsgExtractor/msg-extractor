"""
Debug variable used throughout the entire module.
Turns on debugging information.
"""

import json
import logging
import logging.config
import os
import sys



logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

_debug = False

def getContFileDir(_file_):
    """
    Takes in the path to a file and tries to return the containing folder.
    """
    return '/'.join(_file_.replace('\\', '/').split('/')[:-1])

def toggle_debug():
    global _debug
    _debug = bool(_debug ^ 1)
    if _debug:
        logging.getLogger().setLevel(logging.DEBUG)
    else:
        logging.getLogger().setLevel(logging.WARNING)

def activate_debug():
    global _debug
    _debug = True
    logging.getLogger().setLevel(logging.DEBUG)

def setup_logging(default_path=None, default_level=logging.WARN, env_key='EXTRACT_MSG_LOG_CFG'):
    """Setup logging configuration

    Args:
        default_path (str): Default path to use for the logging configuration file
        default_level (int): Default logging level
        env_key (str): Environment variable name to search for, for setting logfile path

    Returns:
        bool: True if the configuration file was found and applied, False otherwise
    """
    shipped_config = getContFileDir(__file__) + '/logging-config/'
    if os.name == 'nt':
        shipped_config += 'logging-nt.json'
    elif os.name == 'posix':
        shipped_config += 'logging-posix.json'
    # Find logging.json if not provided
    if not default_path:
        default_path = shipped_config

    paths = [
        default_path,
        'logging.json',
        '../logging.json',
        '../../logging.json',
        shipped_config,
    ]

    path = None

    for config_path in paths:
        if os.path.exists(config_path):
            path = config_path
            break

    value = os.getenv(env_key, None)
    if value and os.path.exists(value):
        path = value

    if path is None:
        print('Unable to find logging.json configuration file')
        print('Make sure a valid logging configuration file is referenced in the default_path'
              ' argument, is inside the extract_msg install location, or is available at one '
              'of the following file-paths:')
        print(str(paths[1:]))
        logging.basicConfig(level=default_level)
        logger.warning('The extract_msg logging configuration was not found - using a basic configuration.'
                       'Please check the extract_msg installation directory for "logging-{}.json".'.format(os.name))
        return False

    with open(path, 'rt') as f:
        config = json.load(f)

    try:
        for x in config['handlers']:
            if 'filename' in config['handlers'][x]:
                config['handlers'][x]['filename'] = tmp = os.path.expanduser(os.path.expandvars(config['handlers'][x]['filename']))
                tmp = getContFileDir(tmp)
                if not os.path.exists(tmp):
                    os.makedirs(tmp)
    except:
        print('except')
        pass
    try:
        logging.config.dictConfig(config)
    except ValueError as e:
        print('Failed to configure the logger. Did your installation get messed up?')
        print(e)

    return True
