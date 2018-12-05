|License: GPL v3| |PyPI3| |PyPI1| |PyPI2|

msg-extractor
=============

Extracts emails and attachments saved in Microsoft Outlook's .msg files

The python package extract_msg automates the extraction of key email
data (from, to, cc, date, subject, body) and the email's attachments.

-  `Changelog`_

Usage
-----

**To use it as a command-line script**:

::

     python extract_msg example.msg

This will produce a new folder named according to the date, time and
subject of the message (for example "2013-07-24_0915 Example"). The
email itself can be found inside the new folder along with the
attachments. As of version 0.2, it is capable of extracting both ASCII
and Unicode data.

The script uses Philippe Lagadec's Python module that reads Microsoft
OLE2 files (also called Structured Storage, Compound File Binary Format
or Compound Document File Format). This is the underlying format of
Outlook's .msg files. This library currently supports up to Python 2.7
and 3.4.

The script was built using Peter Fiskerstrand's documentation of the
.msg format. Redemption's discussion of the different property types
used within Extended MAPI was also useful. For future reference, I note
that Microsoft have opened up their documentation of the file format.

If you are having difficulty with a specific file, or would like to
extract more than is currently automated, then the --raw flag may be
useful:

::

     python extract_msg --raw example.msg

Further, a --json flag has been added by Joel Kaufman to specify JSON
output:

::

     python extract_msg --json example.msg

Joel also added a --use-file-name flag, which allows you to specify that
the script writes the emails' contents to the names of the .msg files,
rather than using the subject and date to name the folder:

::

     python extract_msg --use-file-name example.msg

Creation also added a --use-content-id flag, which allows you to specify
that attachments should be saved under the name of their content id,
should they have one. This can be useful for matching attachments to the
names used in the HTML body, and can be done like so:

::

     python extract_msg --use-content-id example.msg

**To use this in your own script**, start by using:

::

     import extract_msg

From there, initialize an instance of the Message class:

::

    

.. _Changelog: CHANGELOG.md

.. |License: GPL v3| image:: https://img.shields.io/badge/License-GPLv3-blue.svg
   :target: LICENSE.txt
.. |PyPI3| image:: https://img.shields.io/badge/pypi-0.20.8-blue.svg
   :target: https://pypi.org/project/extract-msg/0.20.8/
.. |PyPI1| image:: https://img.shields.io/badge/python-2.7+-brightgreen.svg
   :target: https://www.python.org/downloads/release/python-2715/
.. |PyPI2| image:: https://img.shields.io/badge/python-3.6+-brightgreen.svg
   :target: https://www.python.org/downloads/release/python-367/