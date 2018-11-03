msg-extractor
=============

Extracts emails and attachments saved in Microsoft Outlook's .msg files

The python script ExtractMsg.py automates the extraction of key email data (from, to, cc, date, subject, body) and the email's attachments.

Usage
------------

**To use it as a command-line script**:
```
  python ExtractMsg.py example.msg
```

This will produce a new folder named according to the date, time and subject of the message (for example "2013-07-24_0915 Example").  The email itself can be found inside the new folder along with the attachments.  As of version 0.2, it is capable of extracting both ASCII and Unicode data.

The script uses <a href="http://www.decalage.info/python/olefileio">Philippe Lagadec's Python module</a> that reads Microsoft OLE2 files (also called Structured Storage, Compound File Binary Format or Compound Document File Format).  This is the underlying format of Outlook's .msg files.  This library currently supports up to Python 2.7 and 3.4. 

The script was built using <a href="http://www.fileformat.info/format/outlookmsg/index.htm">Peter Fiskerstrand's documentation of the .msg format</a>.  <a href="http://www.dimastr.com/redemption/utils.htm">Redemption's discussion of the different property types used within Extended MAPI</a> was also useful.  For future reference, I note that Microsoft have opened up <a href="http://msdn.microsoft.com/en-us/library/cc463912%28v=exchg.80%29.aspx">their documentation of the file format</a>.

If you are having difficulty with a specific file, or would like to extract more than is currently automated, then the --raw flag may be useful:
```
  python ExtractMsg.py --raw example.msg
```

Further, a --json flag has been added by Joel Kaufman to specify JSON output:
```
  python ExtractMsg.py --json example.msg
```

Joel also added a --use-file-name flag, which allows you to specify that the script writes the emails' contents to the names of the .msg files, rather than using the subject and date to name the folder:
```
  python ExtractMsg.py --use-file-name example.msg
```

Creation also added a --use-content-id flag, which allows you to specify that attachments should be saved under the name of their content id, should they have one.  This can be useful for mathich attachments to the names used in the HTML body, and can be done like so:
```
  python ExtractMsg.py --use-content-id example.msg
```

**To use this in your own script**, start by using:
```
  import ExtractMsg
```

From there, initialize an instance of the Message class:
```
  msg = ExtractMsg.Message("path/to/msg/file.msg")
```
or
```
  msg = ExtractMsg.Message(msg_data_stream)
```

If you want to override the default attachment class and use one of your own, simply change the code to:
```
  msg = ExtractMsg.Message("path/to/msg/file.msg", attachmentClass = CustomAttachmentClass)
```
where `CustomAttachmentClass` is your custom class.

#TODO: Finish this section


If you have any questions feel free to contact me, Matthew Walker, at mattgwwalker at gmail.com.
NOTE: Due to time constraints, <a href="https://github.com/TheElementalOfCreation">The Elemental of Creation</a> has been added as a contributer to help manage the project.  As such, it may be helpful to send emails to arceusthe@gmail.com as well.

Error Reporting
------------
Should you encounter an error that has not already been reported, please do the following when reporting it:
* Make sure you are using the latest version of ExtractMsg.
* State your Python version.
* Include the code, if any, that you used.
* Include a copy of the traceback.

Installation
------------

You can install using pip with:
```sh
  pip install https://github.com/mattgwwalker/msg-extractor/zipball/master
```

or you can include this in your list of python dependencies with:
```python
# setup.py

setup(
    ...
    dependency_links=['https://github.com/mattgwwalker/msg-extractor/zipball/master'],
)
```
