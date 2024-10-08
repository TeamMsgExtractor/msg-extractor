__all__ = [
    'TEST_FILE_DIR',
    'USER_TEST_DIR',
]


import os

from pathlib import Path


TEST_FILE_DIR = Path(__file__).parent.parent / 'example-msg-files'
_utd = None
if bool(userTestDir := os.environ.get('EXTRACT_MSG_TEST_DIR')):
    userTestDir = Path(userTestDir)
    if userTestDir.exists():
        _utd = userTestDir

USER_TEST_DIR = _utd
