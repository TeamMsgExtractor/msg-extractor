import logging
import sys
import traceback

from extract_msg import __doc__, utils
from extract_msg.compat import os_ as os
from extract_msg.message import Message

def main():
    # Setup logging to stdout, indicate running from cli
    CLI_LOGGING = 'extract_msg_cli'

    args = utils.get_command_args(sys.argv[1:])
    level = logging.INFO if args.verbose else logging.WARNING
    currentdir = os.getcwdu() # Store this just in case the paths that have been given are relative
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

        val_results = {x[0]: validation.validate(x[0]) for x in args.msgs}
        filename = 'validation {}.json'.format(int(time.time()))
        print('Validation Results:')
        pprint.pprint(val_results)
        print('These results have been saved to {}'.format(filename))
        with open(filename, 'w') as fil:
            fil.write(json.dumps(val_results))
        utils.get_input('Press enter to exit...')
    else:
        if not args.dump_stdout:
            utils.setup_logging(args.config_path, level, args.log, args.file_logging)
        for x in args.msgs:
            try:
                with Message(x[0]) as msg:
                    # Right here we should still be in the path in currentdir
                    if args.dump_stdout:
                        print(msg.body)
                    else:
                        os.chdir(out)
                        msg.save(toJson = args.json, useFileName = args.use_filename, ContentId = args.cid)#, html = args.html, rtf = args.html, args.allowFallback)
            except Exception as e:
                print("Error with file '" + x[0] + "': " +
                      traceback.format_exc())
            os.chdir(currentdir)

if __name__ == '__main__':
    main()
