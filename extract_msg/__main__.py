__all__ = [
    'main',
]


import os
import sys
import traceback
import zipfile

from extract_msg import __doc__, openMsg, utils
from extract_msg.enums import ErrorBehavior
from typing import List


def main(argv: List[str] = sys.argv) -> None:
    # Setup logging to stdout, indicate running from cli
    CLI_LOGGING = 'extract_msg_cli'
    args = utils.getCommandArgs(argv[1:])

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

    if not args.dumpStdout:
        utils.setupLogging(args.configPath, args.logLevel, args.log, args.fileLogging)

    if args.zip:
        createdZip = True
        _zip = zipfile.ZipFile(args.zip, 'a', zipfile.ZIP_DEFLATED)
    else:
        createdZip = False
        _zip = None

    # Quickly make a dictionary for the keyword arguments.
    kwargs = {
        'allowFallback': args.allowFallback,
        'attachmentsOnly': args.attachmentsOnly,
        'charset': args.charset,
        'contentId': args.cid,
        'customFilename': args.outName,
        'customPath': out,
        'extractEmbedded': args.extractEmbedded,
        'html': args.html,
        'json': args.json,
        'overwriteExisting': args.overwriteExisting,
        'pdf': args.pdf,
        'preparedHtml': args.preparedHtml,
        'rtf': args.rtf,
        'saveHeader': args.saveHeader,
        'skipBodyNotFound': args.skipBodyNotFound,
        'skipEmbedded': args.skipEmbedded,
        'skipHidden': args.skipHidden,
        'skipNotImplemented': args.skipNotImplemented,
        'useMsgFilename': args.useFilename,
        'wkOptions': args.wkOptions,
        'wkPath': args.wkPath,
        'zip': _zip,
    }

    openKwargs = {
        'errorBehavior': ErrorBehavior.RTFDE if args.ignoreRtfDeErrors else ErrorBehavior.THROW,
    }

    # If we are skipping the NotImplementedError attachments, we need to
    # suppress the error.
    if args.skipNotImplemented:
        openKwargs['errorBehavior'] |= ErrorBehavior.ATTACH_NOT_IMPLEMENTED

    def strSanitize(inp):
        """
        Small function to santize parts of a string when failing to print
        them.
        """
        return ''.join((x if x.isascii() else
                f'\\x{ord(x):02X}' if ord(x) <= 0xFF else
                f'\\u{ord(x):04X}' if ord(x) <= 0xFFFF else
                f'\\U{ord(x):08X}') for x in repr(inp))

    for x in args.msgs:
        if args.progress:
            # This may throw an error sometimes and not othertimes.
            # Unclear why, so let's just silence it.
            try:
                print(f'Saving file "{x}"...')
            except UnicodeEncodeError:
                print(f'Saving file "{strSanitize(x)}" (failed to print without repr)...')
        try:
            with openMsg(x, **openKwargs) as msg:
                if args.dumpStdout:
                    print(msg.body)
                elif args.noFolders:
                    msg.saveAttachments(**kwargs)
                else:
                    msg.save(**kwargs)
        except Exception as e:
            try:
                print(f'Error with file "{x}": {traceback.format_exc()}')
            except UnicodeEncodeError:
                print(f'Error with file "{strSanitize(x)}": {traceback.format_exc()}')

    # Close the zip file if we opened it.
    if createdZip:
        _zip.close()

if __name__ == '__main__':
    main(sys.argv)
