msg-extractor
=============

Extracts emails and attachments saved in Microsoft Outlook's .msg files

The python package extract_msg automates the extraction of key email data (from, to, cc, date, subject, body) and the email's attachments.

Usage
-----

**To use it as a command-line script**:
```
  python extract_msg example.msg
```

This will produce a new folder named according to the date, time and subject of the message (for example "2013-07-24_0915 Example").  The email itself can be found inside the new folder along with the attachments.  As of version 0.2, it is capable of extracting both ASCII and Unicode data.

The script uses <a href="http://www.decalage.info/python/olefileio">Philippe Lagadec's Python module</a> that reads Microsoft OLE2 files (also called Structured Storage, Compound File Binary Format or Compound Document File Format).  This is the underlying format of Outlook's .msg files.  This library currently supports up to Python 2.7 and 3.4.

The script was built using <a href="http://www.fileformat.info/format/outlookmsg/index.htm">Peter Fiskerstrand's documentation of the .msg format</a>.  <a href="http://www.dimastr.com/redemption/utils.htm">Redemption's discussion of the different property types used within Extended MAPI</a> was also useful.  For future reference, I note that Microsoft have opened up <a href="http://msdn.microsoft.com/en-us/library/cc463912%28v=exchg.80%29.aspx">their documentation of the file format</a>.

If you are having difficulty with a specific file, or would like to extract more than is currently automated, then the --raw flag may be useful:
```
  python extract_msg --raw example.msg
```

Further, a --json flag has been added by Joel Kaufman to specify JSON output:
```
  python extract_msg --json example.msg
```

Joel also added a --use-file-name flag, which allows you to specify that the script writes the emails' contents to the names of the .msg files, rather than using the subject and date to name the folder:
```
  python extract_msg --use-file-name example.msg
```

Creation also added a --use-content-id flag, which allows you to specify that attachments should be saved under the name of their content id, should they have one.  This can be useful for matching attachments to the names used in the HTML body, and can be done like so:
```
  python extract_msg --use-content-id example.msg
```

**To use this in your own script**, start by using:
```
  import extract_msg
```

From there, initialize an instance of the Message class:
```
  msg = extract_msg.Message("path/to/msg/file.msg")
```

Alternatively, if you wish to send a msg binary string instead of a file to the ExtractMsg.Message Method:
```
  msg_raw = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1\x0 ... \x00x00x00'
  msg = extract_msg.Message(msg_raw)
```

If you want to override the default attachment class and use one of your own, simply change the code to:
```
  msg = extract_msg.Message("path/to/msg/file.msg", attachmentClass = CustomAttachmentClass)
```
where `CustomAttachmentClass` is your custom class.

#TODO: Finish this section


If you have any questions feel free to contact me, Matthew Walker, at mattgwwalker at gmail.com.
NOTE: Due to time constraints, <a href="https://github.com/TheElementalOfCreation">The Elemental of Creation</a> has been added as a contributor to help manage the project.  As such, it may be helpful to send emails to arceusthe@gmail.com as well.

If you have issues, it would be best to get help for them by opening a new github issue.

Error Reporting
------------
Should you encounter an error that has not already been reported, please do the following when reporting it:
* Make sure you are using the latest version of extract_msg.
* State your Python version.
* Include the code, if any, that you used.
* Include a copy of the traceback.

Installation
------------

You can install using pip with:
```sh
  pip install git+https://github.com/mattgwwalker/msg-extractor
```

or you can include this in your list of python dependencies with:
```python
# setup.py

setup(
    ...
    dependency_links=['https://github.com/mattgwwalker/msg-extractor/zipball/master'],
)
```

Todo
------------
Here is a list of things that are currently on our todo list:
* Tests (ie. unittest)
* Finish writing a usage guide
* Improve the intelligence of the saving functions
* Create a Pypi package
* Provide way to save attachments and messages into a custom location under a custom name
* Implement better property handling that will convert each type into a python equivelent if possible
* Implement handling of named properties
