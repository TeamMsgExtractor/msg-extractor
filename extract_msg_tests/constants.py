__all__ = [
    'TEST_FILE_DIR',
    'USER_TEST_DIR',
]


import os

from pathlib import Path


TEST_FILE_DIR = Path(__file__).parent.parent / 'example-msg-files'
USER_TEST_DIR = None
if bool(userTestDir := os.environ.get('EXTRACT_MSG_TEST_DIR')):
    userTestDir = Path(userTestDir)
    if userTestDir.exists():
        USER_TEST_DIR = userTestDir # type: ignore
