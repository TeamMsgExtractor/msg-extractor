import logging
import os
import sys
import traceback

from extract_msg import __doc__, utils
from extract_msg.message import Message

def main():
    # Setup logging to stdout, indicate running from cli
    CLI_LOGGING = 'extract_msg_cli'

    args = utils.getCommandArgs(sys.argv[1:])
    level = logging.INFO if args.verbose else logging.WARNING
    currentdir = os.getcwd() # Store this just in case the paths that have been given are relative
    if args.out_path:
        if not os.path.exists(args.out_path):
            os.makedirs(args.out_path)
        out = args.out_path
    else:
        out = currentdir
    if args.dev:
        import extract_msg.dev
        extract_msg.dev.main(args, sys.argv[1:])
    elif args.validate:
        import json
        import pprint
        import time

        from extract_msg import validation

        valResults = {x[0]: validation.validate(x[0]) for x in args.msgs}
        filename = f'validation {int(time.time())}.json'
        print('Validation Results:')
        pprint.pprint(valResults)
        print(f'These results have been saved to {filename}')
        with open(filename, 'w') as fil:
            fil.write(json.dumps(valResults))
        input('Press enter to exit...')
    else:
        if not args.dump_stdout:
            utils.setupLogging(args.config_path, level, args.log, args.file_logging)
        for x in args.msgs:
            try:
                with Message(x[0]) as msg:
                    # Right here we should still be in the path in currentdir
                    if args.dump_stdout:
                        print(msg.body)
                    else:
                        msg.save(customPath = out, customFilename = args.out_name, json = args.json, useMsgFilename = args.use_filename, contentId = args.cid, html = args.html, rtf = args.rtf, allowFallback = args.allowFallback)
            except Exception as e:
                print("Error with file '" + x[0] + "': " +
                      traceback.format_exc())

if __name__ == '__main__':
    main()
