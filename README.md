msg-extractor
=============

Extracts emails and attachments saved in Microsoft Outlook's .msg files

This python script automates the extraction of key email data (from, to, cc, date, subject, body) and the email's attachments.

It uses Philippe Lagadec's Python module [1] that reads Microsoft OLE2 files (also called Structured Storage, Compound File Binary Format or Compound Document File Format).  This is the underlying format of Outlook's .msg files.

The script also makes the most of Peter Fiskerstrand's documentation of the .msg format. [2]




[1] http://www.decalage.info/python/olefileio
[2] http://www.fileformat.info/format/outlookmsg/index.htm
