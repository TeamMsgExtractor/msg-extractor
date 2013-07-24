msg-extractor
=============

Extracts emails and attachments saved in Microsoft Outlook's .msg files

The python script ExtractMsg.py automates the extraction of key email data (from, to, cc, date, subject, body) and the email's attachments.

To use it
```
  python ExtractMsg.py example.msg
```

This will produce a new folder named according to the date, time and subject of the message (for example "2013-07-24_0915 Example").  The email itself can be found inside the new folder along with the attachments.

The script uses <a href="http://www.decalage.info/python/olefileio">Philippe Lagadec's Python module</a> that reads Microsoft OLE2 files (also called Structured Storage, Compound File Binary Format or Compound Document File Format).  This is the underlying format of Outlook's .msg files.

The script was built using <a href="http://www.fileformat.info/format/outlookmsg/index.htm">Peter Fiskerstrand's documentation of the .msg format</a>.

If you have any questions feel free to contact me, Matthew Walker, at mattgwwalker at gmail.com.
