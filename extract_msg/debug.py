"""
Debug variable used throughout the entire module.
Turns on debugging information.
"""

import logging
import os
from extract_msg.utils import getPyFileDir



logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

_debug = False

def toggle_debug():
    global _debug
    _debug = bool(_debug ^ 1)
    if _debug:
        logger.setLevel(logging.DEBUG)
    else:
        logger.setLevel(logging.WARNING)

def setup_logging(default_path=None, default_level=logging.INFO, env_key='LOG_CFG'):
    """Setup logging configuration

    Args:
        default_path (str): Default path to use for the logging configuration file
        default_level (int): Default logging level
        env_key (str): Environment variable name to search for, for setting logfile path

    Returns:
        bool: True if the configuration file was found and applied, False otherwise
    """
    shipped_config = getPyFileDir(__file__) + '/logging-config/'
    if os.name == 'nt':
        shipped_config += 'logging-nt.json'
    elif os.name == 'posix':
        shipped_config += 'logging-posix.json'
    # Find logging.json if not provided
    if not default_path:
        default_path = getPyFileDir(__file__) + '/logging.json'

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

    if path is None:
        print('Unable to find logging.json configuration file')
        print('Make sure a valid logging configuration file is referenced in the default_path argument or available at '
              'one of the following file-paths:')
        print(str(paths[1:]))
        logging.basicConfig(level=default_level)
        logger.warning('The extract_msg logging configuration was not found - using a basic configuration.'
                       'Please check the integrations directory for "logging.json".')
        return False

    value = os.getenv(env_key, None)
    if value and os.path.exists(value):
        path = value

    with open(path, 'rt') as f:
        config = json.load(f)

    try:
        logging.config.dictConfig(config)
    except ValueError as e:
        print('Failed to find file - make sure your integrations log file is properly configured')
        print(e)

    return True
