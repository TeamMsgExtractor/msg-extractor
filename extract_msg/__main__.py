import glob
import logging
import sys
import traceback
from extract_msg import __doc__
from extract_msg.message import Message

if __name__ == '__main__':

    # command_args = get_cmd_args()

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
        sys.exit(1)

    # Setup logging to stdout, indicate running from cli
    CLI_LOGGING = 'extract_msg_cli'

    writeRaw = False
    toJson = False
    useFileName = False
    useContentId = False
    debug = False

    for rawFilename in sys.argv[1:]:
        if rawFilename == '--raw':
            writeRaw = True

        if rawFilename == '--json':
            toJson = True

        if rawFilename == '--use-file-name':
            useFileName = True

        if rawFilename == '--use-content-id':
            useContentId = True

        if rawFilename == '--use-logging-json':
            log_config = '../logging.json'

        if rawFilename == '--debug':
            debug = True
            logging.getLogger(CLI_LOGGING).setLevel(logging.DEBUG)
        elif rawFilename == '--verbose':
            logging.getLogger(CLI_LOGGING).setLevel(logging.INFO)
        else:
            logging.getLogger(CLI_LOGGING).setLevel(logging.WARNING)

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
