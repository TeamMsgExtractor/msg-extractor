"""
extract_msg:
    Extracts emails and attachments saved in Microsoft Outlook's .msg files

https://github.com/mattgwwalker/msg-extractor
"""

__author__ = 'Matthew Walker & The Elemental of Creation'
__date__ = '2018-05-22'
__version__ = '0.20.2'

import glob
import sys
import traceback
from extract_msg.base import Attachment, Properties, Props, Recipient, Message, msg_epoch, parse_type, properHex
from extract_msg import constants

if __name__ == '__main__':
    if len(sys.argv) <= 1:
        print(__doc__)
        print("""
Launched from command line, this script parses Microsoft Outlook Message files
and save their contents to the current directory. On error the script will
write out a 'raw' directory will all the details from the file, but in a
less-than-desirable format. To force this mode, the flag '--raw'
can be specified.

Usage:  <file> [file2 ...]
   or:  --raw <file>
   or:  --json

Additionally, use the flag '--use-content-id' to save files by their content ID (should they have one)

To name the directory as the .msg file, use the flag '--use-file-name'

To turn on the printing of debugging information, use the flag '--debug'
""")
        sys.exit()

    writeRaw = False
    toJson = False
    useFileName = False
    useContentId = False

    for rawFilename in sys.argv[1:]:
        if rawFilename == '--raw':
            writeRaw = True

        if rawFilename == '--json':
            toJson = True

        if rawFilename == '--use-file-name':
            useFileName = True

        if rawFilename == '--use-content-id':
            useContentId = True

        if rawFilename == '--debug':
            debug = True

        for filename in glob.glob(rawFilename):
            msg = Message(filename)
            try:
                if writeRaw:
                    msg.saveRaw()
                else:
                    msg.save(toJson, useFileName, False, useContentId)
            except Exception as e:
                # msg.debug()
                print("Error with file '" + filename + "': " +
                      traceback.format_exc())
