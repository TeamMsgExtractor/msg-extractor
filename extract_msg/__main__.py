import logging
import os
import sys
import traceback

from extract_msg import __doc__, utils


def main() -> None:
    # Setup logging to stdout, indicate running from cli
    CLI_LOGGING = 'extract_msg_cli'
    args = utils.getCommandArgs(sys.argv[1:])
    level = logging.INFO if args.verbose else logging.WARNING

    # Determine where to save the files to.
    currentDir = os.getcwd() # Store this in case the path changes.
    if not args.zip:
        if args.outPath:
            if not os.path.exists(args.outPath):
                os.makedirs(args.outPath)
            out = args.outPath
        else:
            out = currentDir
    else:
        out = args.outPath if args.outPath else ''

    if args.dev:
        import extract_msg.dev
        extract_msg.dev.main(args, sys.argv[1:])
    elif args.validate:
        import json
        import pprint
        import time

        from extract_msg import validation

        valResults = {x: validation.validate(x) for x in args.msgs}
        filename = f'validation {int(time.time())}.json'
        print('Validation Results:')
        pprint.pprint(valResults)
        print(f'These results have been saved to {filename}')
        with open(filename, 'w') as fil:
            json.dump(valResults, fil)
        input('Press enter to exit...')
    else:
        if not args.dumpStdout:
            utils.setupLogging(args.configPath, level, args.log, args.fileLogging)

        # Quickly make a dictionary for the keyword arguments.
        kwargs = {
            'allowFallback': args.allowFallback,
            'attachmentsOnly': args.attachmentsOnly,
            'charset': args.charset,
            'contentId': args.cid,
            'customFilename': args.outName,
            'customPath': out,
            'html': args.html,
            'json': args.json,
            'pdf': args.pdf,
            'preparedHtml': args.preparedHtml,
            'rtf': args.rtf,
            'useMsgFilename': args.useFilename,
            'wkOptions': args.wkOptions,
            'wkPath': args.wkPath,
            'zip': args.zip,
        }

        openKwargs = {
            'ignoreRtfDeErrors': args.ignoreRtfDeErrors,
        }

        for x in args.msgs:
            if args.progress:
                # This may throw an error sometimes and not othertimes.
                # Unclear why, so let's just silence it.
                try:
                    print(f'Saving file "{x}"...')
                except UnicodeEncodeError:
                    print(f'Saving file "{repr(x)}" (failed to print without repr)...')
            try:
                with utils.openMsg(x, **openKwargs) as msg:
                    if args.dumpStdout:
                        print(msg.body)
                    else:
                        msg.save(**kwargs)
            except Exception as e:
                try:
                    print(f'Error with file "{x}": {traceback.format_exc()}')
                except UnicodeEncodeError:
                    print(f'Error with file "{repr(x)}": {traceback.format_exc()}')


if __name__ == '__main__':
    main()
