|License: GPL v3| |PyPI3| |PyPI2|

msg-extractor
=============

Extracts emails and attachments saved in Microsoft Outlook's .msg files

The python package extract_msg automates the extraction of key email
data (from, to, cc, date, subject, body) and the email's attachments.

NOTICE
======
0.29.* is the branch that supports both Python 2 and Python 3. It is now only
receiving bug fixes and will not be receiving feature updates.

This module has a Discord server for general discussion. You can find it here:
`Discord`_


Changelog
---------
-  `Changelog <CHANGELOG.md>`__

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
Outlook's .msg files. This library currently supports Python 3.6 and above.

The script was originally built using Peter Fiskerstrand's documentation of the
.msg format. Redemption's discussion of the different property types used within
Extended MAPI was also useful. For future reference, I note that Microsoft have
opened up their documentation of the file format, which is what is currently
being used for development.


#########REWRITE COMMAND LINE USAGE#############
Currently, the README is in the process of being redone. For now, please
refer to the usage information provided from the program's help dialog:
::

    usage: extract_msg [-h] [--use-content-id] [--dev] [--validate] [--json]
                       [--file-logging] [--verbose] [--log LOG]
                       [--config CONFIG_PATH] [--out OUT_PATH] [--use-filename]
                       [--dump-stdout] [--html] [--raw] [--rtf]
                       [--allow-fallback] [--out-name OUT_NAME] msg [msg ...]

    extract_msg: Extracts emails and attachments saved in Microsoft Outlook's
    .msg files. https://github.com/TeamMsgExtractor/msg-extractor

    positional arguments:
        msg                   An msg file to be parsed

    optional arguments:
        -h, --help            show this help message and exit
        --use-content-id, --cid
                              Save attachments by their Content ID, if they have
                              one. Useful when working with the HTML body.
        --dev                 Changes to use developer mode. Automatically
                              enables the --verbose flag. Takes precedence over
                              the --validate flag.
        --validate            Turns on file validation mode. Turns off regular
                              file output.
        --json                Changes to write output files as json.
        --file-logging        Enables file logging. Implies --verbose.
        --verbose             Turns on console logging.
        --log LOG             Set the path to write the file log to.
        --config CONFIG_PATH  Set the path to load the logging config from.
        --out OUT_PATH        Set the folder to use for the program output.
                              (Default: Current directory)
        --use-filename        Sets whether the name of each output is based on
                              the msg filename.
        --dump-stdout         Tells the program to dump the message body (plain
                              text) to stdout. Overrides saving arguments.
        --html                Sets whether the output should be html. If this is
                              not possible, will error.
        --raw                 Sets whether the output should be raw. If this is
                              not possible, will error.
        --rtf                 Sets whether the output should be rtf. If this is
                              not possible, will error.
        --allow-fallback      Tells the program to fallback to a different save
                              type if the selected one is not possible.
        --out-name OUT_NAME   Name to be used with saving the file output.
                              Should come immediately after the file name.

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

If you have any questions feel free to contact me, Destiny, as arceusthe [at]
gmail [dot] com. I am the co-owner and main developer of the project.

If you have issues, it would be best to get help for them by opening a
new github issue.

Error Reporting
---------------

Should you encounter an error that has not already been reported, please
do the following when reporting it: \* Make sure you are using the
latest version of extract_msg (check the version on PyPi). \* State your
Python version. \* Include the code, if any, that you used. \* Include a
copy of the traceback.

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

Todo
----

Here is a list of things that are currently on our todo list:

* Tests (ie. unittest)
* Finish writing a usage guide
* Improve the intelligence of the saving functions
* Improve handling of named properties
* Improve README
* Create a wiki for advanced usage information

Credits
-------

`Matthew Walker`_ - Original developer and owner.

`Destiny Peterson (The Elemental of Destruction)`_ - Co-owner, principle programmer, knows more about msg files than anyone probably should.

`JP Bourget`_ - Senior programmer, readability and organization expert, secondary manager.

`Philippe Lagadec`_ - Python OleFile module developer.

`Joel Kaufman`_ - First implementations of the json and filename flags.

`Dean Malmgren`_ - First implementation of the setup.py script.

`Seamus Tuohy`_ - Developer of the Python RTFDE module. Gave first examples of how to use the module.

`Liam`_ - Significant reorganization and transfer of data.

And thank you to everyone who has opened an issue and helped us track down those pesky bugs.

.. |License: GPL v3| image:: https://img.shields.io/badge/License-GPLv3-blue.svg
   :target: LICENSE.txt

.. |PyPI3| image:: https://img.shields.io/badge/pypi-0.30.13-blue.svg
   :target: https://pypi.org/project/extract-msg/0.30.13/

.. |PyPI2| image:: https://img.shields.io/badge/python-3.6+-brightgreen.svg
   :target: https://www.python.org/downloads/release/python-367/
.. _Matthew Walker: https://github.com/mattgwwalker
.. _Destiny Peterson (The Elemental of Destruction): https://github.com/TheElementalOfDestruction
.. _JP Bourget: https://github.com/punkrokk
.. _Philippe Lagadec: https://github.com/decalage2
.. _Dean Malmgren: https://github.com/deanmalmgren
.. _Joel Kaufman: https://github.com/joelkaufman
.. _Liam: https://github.com/LiamPM5
.. _Seamus Tuohy: https://github.com/seamustuohy
.. _Discord: https://discord.com/invite/B77McRmzdc
