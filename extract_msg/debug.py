"""
Debug variable used throughout the entire module.
Turns on debugging information.
"""

import json
import logging
import logging.config
import os

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
    # Find logging.json if not provided
    if default_path:
        default_path = default_path

    paths = [
        default_path,
        '../logging.json',
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
        logger.warning('See example logging.json examples in the templates folder.')
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
