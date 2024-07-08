|License: GPL v3| |PyPI3| |PyPI2| |Read the Docs|

extract-msg
=============

Extracts emails and attachments saved in Microsoft Outlook's .msg files

The python package extract_msg automates the extraction of key email
data (from, to, cc, date, subject, body) and the email's attachments.

Documentation can be found in the code, on the `wiki`_, and on the
`Read the Docs`_ page.

NOTICE
======
0.29.* is the branch that supports both Python 2 and Python 3. It is now only
receiving bug fixes and will not be receiving feature updates.

0.39.* is the last versions that supported Python 3.6 and 3.7. Support for those
was dropped to allow the use of new features from 3.8 and because the life spans
of those versions had ended.

This module has a Discord server for general discussion. You can find it here:
`Discord`_


Changelog
---------
-  `Changelog`_

Usage
-----

**To use it as a command-line script**:

::

     python -m extract_msg example.msg

This will produce a new folder named according to the date, time and
subject of the message (for example "2013-07-24_0915 Example"). The
email itself can be found inside the new folder along with the
attachments.

The script uses Philippe Lagadec's Python module that reads Microsoft
OLE2 files (also called Structured Storage, Compound File Binary Format
or Compound Document File Format). This is the underlying format of
Outlook's .msg files. This library currently supports Python 3.8 and above.

The script was originally built using Peter Fiskerstrand's documentation of the
.msg format. Redemption's discussion of the different property types used within
Extended MAPI was also useful. For future reference, note that Microsoft have
opened up their documentation of the file format, which is what is currently
being used for development.


#########REWRITE COMMAND LINE USAGE#############
Currently, the README is in the process of being redone. For now, please
refer to the usage information provided from the program's help dialog:
::

     usage: extract_msg [-h] [--use-content-id] [--json] [--file-logging] [-v] [--log LOG] [--config CONFIGPATH] [--out OUTPATH] [--use-filename] [--dump-stdout] [--html] [--pdf] [--wk-path WKPATH] [--wk-options [WKOPTIONS ...]]
                        [--prepared-html] [--charset CHARSET] [--raw] [--rtf] [--allow-fallback] [--skip-body-not-found] [--zip ZIP] [--save-header] [--attachments-only] [--skip-hidden] [--no-folders] [--skip-embedded] [--extract-embedded]
                        [--overwrite-existing] [--skip-not-implemented] [--out-name OUTNAME | --glob] [--ignore-rtfde] [--progress] [-s]
                        msg [msg ...]

     extract_msg: Extracts emails and attachments saved in Microsoft Outlook's .msg files. https://github.com/TeamMsgExtractor/msg-extractor

     positional arguments:
       msg                   An MSG file to be parsed.

     options:
       -h, --help            show this help message and exit
       --use-content-id, --cid
                             Save attachments by their Content ID, if they have one. Useful when working with the HTML body.
       --json                Changes to write output files as json.
       --file-logging        Enables file logging. Implies --verbose level 1.
       -v, --verbose         Turns on console logging. Specify more than once for higher verbosity.
       --log LOG             Set the path to write the file log to.
       --config CONFIGPATH   Set the path to load the logging config from.
       --out OUTPATH         Set the folder to use for the program output. (Default: Current directory)
       --use-filename        Sets whether the name of each output is based on the msg filename.
       --dump-stdout         Tells the program to dump the message body (plain text) to stdout. Overrides saving arguments.
       --html                Sets whether the output should be HTML. If this is not possible, will error.
       --pdf                 Saves the body as a PDF. If this is not possible, will error.
       --wk-path WKPATH      Overrides the path for finding wkhtmltopdf.
       --wk-options [WKOPTIONS ...]
                             Sets additional options to be used in wkhtmltopdf. Should be a series of options and values, replacing the - or -- in the beginning with + or ++, respectively. For example: --wk-options "+O Landscape"
       --prepared-html       When used in conjunction with --html, sets whether the HTML output should be prepared for embedded attachments.
       --charset CHARSET     Character set to use for the prepared HTML in the added tag. (Default: utf-8)
       --raw                 Sets whether the output should be raw. If this is not possible, will error.
       --rtf                 Sets whether the output should be RTF. If this is not possible, will error.
       --allow-fallback      Tells the program to fallback to a different save type if the selected one is not possible.
       --skip-body-not-found
                             Skips saving the body if the body cannot be found, rather than throwing an error.
       --zip ZIP             Path to use for saving to a zip file.
       --save-header         Store the header in a separate file.
       --attachments-only    Specify to only save attachments from an msg file.
       --skip-hidden         Skips any attachment marked as hidden (usually ones embedded in the body).
       --no-folders          Stores everything in the location specified by --out. Requires --attachments-only and is incompatible with --out-name.
       --skip-embedded       Skips all embedded MSG files when saving attachments.
       --extract-embedded    Extracts the embedded MSG files as MSG files instead of running their save functions.
       --overwrite-existing  Disables filename conflict resolution code for attachments when saving a file, causing files to be overwriten if two attachments with the same filename are on an MSG file.
       --skip-not-implemented, --skip-ni
                             Skips any attachments that are not implemented, allowing saving of the rest of the message.
       --out-name OUTNAME    Name to be used with saving the file output. Cannot be used if you are saving more than one file.
       --glob, --wildcard    Interpret all paths as having wildcards. Incompatible with --out-name.
       --ignore-rtfde        Ignores all errors thrown from RTFDE when trying to save. Useful for allowing fallback to continue when an exception happens.
       --progress            Shows what file the program is currently working on during it's progress.
       -s, --stdin           Read file from stdin (only works with one file at a time).

**To use this in your own script**, start by using:

::

     import extract_msg

From there, open the MSG file:

::

     msg = extract_msg.openMsg("path/to/msg/file.msg")

Alternatively, if you wish to send a msg binary string instead of a file
to the ``extract_msg.openMsg`` Method:

::

     msg_raw = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1\x00 ... \x00\x00\x00'
     msg = extract_msg.openMsg(msg_raw)

If you want to override the default attachment class and use one of your
own, simply change the code to:

::

     msg = extract_msg.openMsg("path/to/msg/file.msg", attachmentClass = CustomAttachmentClass)

where ``CustomAttachmentClass`` is your custom class.

#TODO: Finish this section

If you have any questions feel free to contact Destiny at arceusthe [at]
gmail [dot] com. She is the co-owner and main developer of the project.

If you have issues, it would be best to get help for them by opening a
new github issue.

Error Reporting
---------------

Should you encounter an error that has not already been reported, please
do the following when reporting it: \* Make sure you are using the
latest version of extract_msg (check the version on PyPi). \* State your
Python version. \* Include the code, if any, that you used. \* Include a
copy of the traceback.

Supporting The Module
---------------------

If you'd like to donate to help support the development of the module, you can
donate to Destiny using one of the following services:

* `Buy Me a Coffee`_
* `Ko-fi`_
* `Patreon`_

Installation
------------

You can install using pip:

-  Pypi

.. code:: bash

       pip install extract-msg

-  Github

.. code:: sh

     pip install git+https://github.com/TeamMsgExtractor/msg-extractor

or you can include this in your list of python dependencies with:

.. code:: python

   # setup.py

   setup(
       ...
       dependency_links=['https://github.com/TeamMsgExtractor/msg-extractor/zipball/master'],
   )

Additionally, this module has the following extras which can be optionally
installed:

* ``all``: Installs all of the extras.
* ``mime``: Installs dependency used for mimetype generation when a mimetype is not specified.

Versioning
----------

This module uses Semantic Versioning, however it has not always done so. All versions greater than or equal to 0.40.* conform successfully. As the package is currently in major version zero (0.*.*), anything MAY change at any time, as per point 4 of the SemVer specification. However, I, Destiny, am aware of the module's usage in other packages and code, and so I have taken efforts to make the versioning more reliable.

Any change to the minor version MUST be considered a potentially breaking change, and the changelog should be checked before assuming the API will function in the way it did in the previous minor version. I do, however, try to keep the API relatively stable between minor versions, so most typical usage is likely to remain entirely unaffected.

Any change to a patch version before the 1.0.0 release SHOULD either add functionality or have no visible difference in usage, aside from changes to the typing infomation or from a bug fix correcting the data that a component created.

In addition to the above conditions, it must be noted that any class, variable, function, etc., that is preceded by one or more underscores, excluding items preceded by two underscores and also proceeded by two underscores, MUST NOT be considered part of the public api. These methods may change at any time, in any way.

I am aware of the F.A.Q. question that suggests that I should probably have pushed the module to a 1.0.0 release due to its usage in production, however there are a number of different items on the TODO list that I feel should be completed before that time. While some are simply important features I believe should exist, others are overhauls to sections of the public API that have needed careful fixing for quite a while, fixes that have slowly been happening throughout the versions. An important change was made in the 0.45.0 release which deprecates a large number of commonly used private functions and created more stable versions of them in the public API.

Additionally, my focus on versioning info has revealed that some of the dependencies are still in major version 0 *or* do not necessarily conform to Semantic Versioning. As such, these packages are more tightly constrained on what versions are considered acceptable, and careful consideration should be taken before extending the accepted range of versions.

Details on Semantic Versioning can be found at `semver.org`_.

Todo
----

Here is a list of things that are currently on our todo list:

* Tests (ie. unittest)
* Finish writing a usage guide
* Improve the intelligence of the saving functions
* Improve README
* Create a wiki for advanced usage information

Credits
-------

`Destiny Peterson (The Elemental of Destruction)`_ - Co-owner, principle programmer, knows more about msg files than anyone probably should.

`Matthew Walker`_ - Original developer and co-owner.

`JP Bourget`_ - Senior programmer, readability and organization expert, secondary manager.

`Philippe Lagadec`_ - Python OleFile module developer.

`Joel Kaufman`_ - First implementations of the json and filename flags.

`Dean Malmgren`_ - First implementation of the setup.py script.

`Seamus Tuohy`_ - Developer of the Python RTFDE module. Gave first examples of how to use the module and has worked with Destiny to ensure functionality.

`Liam`_ - Significant reorganization and transfer of data.

And thank you to everyone who has opened an issue and helped us track down those pesky bugs.

Extra
-----

Check out the new project `msg-explorer`_ that allows you to open MSG files and
explore their contents in a GUI. It is usually updated within a few days of a
major release to ensure continued support. Because of this, it is recommended to
install it to a separate environment (like a vitural env) to not interfere with
your access to the newest major version of extract-msg.

.. |License: GPL v3| image:: https://img.shields.io/badge/License-GPLv3-blue.svg
   :target: LICENSE.txt

.. |PyPI3| image:: https://img.shields.io/badge/pypi-0.48.7-blue.svg
   :target: https://pypi.org/project/extract-msg/0.48.7/

.. |PyPI2| image:: https://img.shields.io/badge/python-3.8+-brightgreen.svg
   :target: https://www.python.org/downloads/release/python-3810/

.. |Read the Docs| image:: https://readthedocs.org/projects/msg-extractor/badge/?version=latest
    :target: https://msg-extractor.readthedocs.io/en/stable/?badge=latest
    :alt: Documentation Status

.. _Matthew Walker: https://github.com/mattgwwalker
.. _Destiny Peterson (The Elemental of Destruction): https://github.com/TheElementalOfDestruction
.. _JP Bourget: https://github.com/punkrokk
.. _Philippe Lagadec: https://github.com/decalage2
.. _Dean Malmgren: https://github.com/deanmalmgren
.. _Joel Kaufman: https://github.com/joelkaufman
.. _Liam: https://github.com/LiamPM5
.. _Seamus Tuohy: https://github.com/seamustuohy
.. _Discord: https://discord.com/invite/B77McRmzdc
.. _Buy Me a Coffee: https://www.buymeacoffee.com/DestructionE
.. _Ko-fi: https://ko-fi.com/destructione
.. _Patreon: https://www.patreon.com/DestructionE
.. _msg-explorer: https://pypi.org/project/msg-explorer/
.. _wiki: https://github.com/TeamMsgExtractor/msg-extractor/wiki
.. _Read the Docs: https://msg-extractor.rtfd.io/
.. _Changelog: https://github.com/TeamMsgExtractor/msg-extractor/blob/master/CHANGELOG.md
.. _`semver.org`: https://semver.org
