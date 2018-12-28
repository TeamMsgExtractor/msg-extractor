"""
Module for collecting data to be sent to the developer.
"""

# NOTE: Order of tasks:
#    1. Check for exceptions:
#        * Check the entire process for exceptions raised by a specific file and log them. If none occur, log something like "No exceptions were detected."
#    2. Run the file through the developer versions of the classes


import logging

from extract_msg import dev_classes
from extract_msg import utils



def setup_dev_logger(default_path=None, logfile = None, env_key='EXTRACT_MSG_LOG_CFG'):
    utils.setup_logging(default_path, 5, logfile, True, env_key)
	



def main(args):
    pass;
