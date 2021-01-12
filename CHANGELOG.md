**v0.28.1**
* [[TeamMsgExtractor #181](https://github.com/TeamMsgExtractor/msg-extractor/issues/181)] Fixed issue in `Attachment` that arose when moving some of the code to a base class.
* Fixed small error in `utils.parse_type` that caused it to incorrectly compare expected and actual length. Fortunately, this had no actual effect aside from a warning.
* Added the `ebcdic` module to the requirements to add more supported encodings.

**v0.28.0**
* [[TeamMsgExtractor #87](https://github.com/TeamMsgExtractor/msg-extractor/issues/87)] Added a new system to handle `NotImplementedError` and other exceptions. All msg classes now have an option called `attachmentErrorBehavior` that tells the class what to do if it has an error. The value should be one of three constants: `ATTACHMENT_ERROR_THROW`, `ATTACHMENT_ERROR_NOT_IMPLEMENTED`, or `ATTACHMENT_ERROR_BROKEN`. `ATTACHMENT_ERROR_THROW` tells the class to not catch and exceptions and just let the user handle them. `ATTACHMENT_ERROR_NOT_IMPLEMENTED` tells the class to catch `NotImplementedError` exceptions and put an instance of `UnsupportedAttachment` in place of a regular attachment. `ATTACHMENT_ERROR_BROKEN` tells the class to catch *all* exceptions and either replace the attachment with `UnsupportedAttachment` if it is a `NotImplementedError` or `BrokenAttachment` for all other exceptions. With both of those options, caught exceptions will be logged.
* In making the previous point work, much code from `Attachment` has been moved to a new class called `AttachmentBase`. Both `BrokenAttachment` and `UnsupportedAttachment` are subclasses of `AttachmentBase` meaning data can be extracted from their streams in the same way as a functioning attachment.
* [[TeamMsgExtractor #162](https://github.com/TeamMsgExtractor/msg-extractor/issues/162)] Pretty sure I actually got it this time. The execution flag should be applied by pip now.
* Fixed typos in some exceptions

**v0.27.16**
* [[TeamMsgExtractor #177](https://github.com/TeamMsgExtractor/msg-extractor/issues/177)] Fixed incorrect struct being used. It should be the correct one now, but further testing will be required to confirm this.
* Fixed log error message in `extract_msg.prop` to actually format a value into the message.

**v0.27.15**
* [[TeamMsgExtractor #177](https://github.com/TeamMsgExtractor/msg-extractor/issues/177)] Fixed missing import.

**v0.27.14**
* [[TeamMsgExtractor #173](https://github.com/TeamMsgExtractor/msg-extractor/issues/173)] Fixed typo that I made in the last version that broke things. I didn't have the resources to test this one myself, unfortunately.
* Fixed a typo in an exception message.

**v0.27.13**
* [[TeamMsgExtractor #173](https://github.com/TeamMsgExtractor/msg-extractor/issues/173)] Moved some data used in checks into constants so that I can make sure they get changed every where that they are used. Hopefully I can close this issue.

**v0.27.12**
* [[TeamMsgExtractor #173](https://github.com/TeamMsgExtractor/msg-extractor/issues/173)] Made an assumption about where an exception was thrown from and was wrong. While that location would have throw an exception, the function that called that code was the one to actually throw the exception in question. This issue *should* be fixed...
* [[TeamMsgExtractor #162](https://github.com/TeamMsgExtractor/msg-extractor/issues/162)] Made another attempt to fix the execution flag on the wrapper script.

**v0.27.11**
* [[TeamMsgExtractor #173](https://github.com/TeamMsgExtractor/msg-extractor/issues/173)] Tentatively implemented type 0x1014 (PtypMultipleInteger64). Apparently I forgot to do it earlier.

**v0.27.10**
* [[TeamMsgExtractor #162](https://github.com/TeamMsgExtractor/msg-extractor/issues/162)] Fixed line endings in the wrapper script to be UNIX line endings rather than Windows line endings. Attempted to add the execution flag to the runnable script.

**v0.27.9**
* [[TeamMsgExtractor #161](https://github.com/TeamMsgExtractor/msg-extractor/issues/161)] Added commands to the command line that will allow the user to specify that they want the message data to be output to stdout rather than to a file.
* [[TeamMsgExtractor #162](https://github.com/TeamMsgExtractor/msg-extractor/issues/162)] Added a wrapper for extract_msg that will be installed.
* Fixed some of the encoding names to allow them to actually be used in Python. The names they previously held were not aliases that currently exist.
* Added more documentation to `constants.CODE_PAGES` to give more information about what it is. As it is a list of the possible encodings an msg file can use, I also specified which ones were supported by Python 3.
* Moved the main code into a function so it is now callable from outside of the file.

**v0.27.8**
* [[TeamMsgExtractor #158](https://github.com/TeamMsgExtractor/msg-extractor/issues/158)] Fixed a spelling error in a function name that was causing it to not be seen. The function was called `ceilDiv` but was accidentally called as `cielDiv`.

**v0.27.7**
* Fixed an issue in the new bitwise adjustment functions. One of the variable names was incorrect.

**v0.27.6**
* Fixed a few lines in `data.py`.

**v0.27.5**
* Fixed an error in `utils.divide` that would cause it to drop the extra data if there was not enough to create a full division. For example, if you had a string that was 10 characters, and divided by 3, you would only receive a total of 9 characters back.
* Added some useful functions that will be used in the future.
* [[TeamMsgExtractor #155](https://github.com/TeamMsgExtractor/msg-extractor/issues/155)] Updated to use new version of tzlocal.
* Updated changelog to fit new repository.

**v0.27.4**
* [[TeamMsgExtractor #152](https://github.com/TeamMsgExtractor/msg-extractor/issues/152)] Fixed an issue where the name of an exception was put as the wrong thing.

**v0.27.3**
* [[TeamMsgExtractor #105](https://github.com/TeamMsgExtractor/msg-extractor/issues/105)] Added code to fix an internal msg issue that had recipient lists being split up in the header after a certain amount of characters. This was now a bug on our part, but an issue with the generation of the msg file itself.
* Exposed the `MessageBase` class directly from `extract_msg`. I forgot to do this when I created it.
* Added `MessageBase.bcc`. I think it used to exist but got erased somehow on accident. Either way, it exists now.

**v0.27.2**
* After much debate, I have finally decided to allow an option to override the string encoding in message files. This Was something I reserved solely for `dev_classes.Message` because it felt like it didn't fit with how msg files were supposed to work. I also didn't want messages from people about them running into errors after they overrode the encoding. You can now do this by providing the `overrideEncoding` option on any `MSGFile` class as well as the `openMsg` function.
* [[TeamMsgExtractor #103](https://github.com/TeamMsgExtractor/msg-extractor/issues/103)] Implemented correct detection of encodings. If you have any more issues with "'X' codec can't decode bytes" it is likely because the encoding specified inside the msg file is wrong.

**v0.27.1**
* [[TeamMsgExtractor #147](https://github.com/TeamMsgExtractor/msg-extractor/issues/147)] Fixed an issue in `Message.save` caused by it attempting to directly access a private variable that was moved to the base class `MessageBase`.

**v0.27.0**
* [[TeamMsgExtractor #143](https://github.com/TeamMsgExtractor/msg-extractor/issues/143)] Added new class `Appointment` that can handle outlook appointments or meetings.
* [[TeamMsgExtractor #143](https://github.com/TeamMsgExtractor/msg-extractor/issues/143)] Added new class `MessageBase` for classes that end up being mostly like the `Message` class. Currently subclassed by `Message` and `Appointment`. This should not have a direct effect on any code that uses this module.
* Added support for `PtypFloatingTime` in `utils.parseType`.
* Added proper support for `PtypTime` in `FixedLengthProp.parseType`
* Added pretty print functions to the `Properties` class and the `Named` class.
* Added new function `MSGFile._ensureSetProperty` that acts like `MSGFile._ensureSet` except that it works with properties from the `Properties` instance.
* Added equivalent of the previously mentioned function to the `Attachment` class and the `Recipient` class.

**v0.26.4**
* Added new function `MSGFile._ensureSetNamed` which acts like `MSGFile._ensureSet` except that it works with named properties.
* Added a version of the previous function to the `Attachment` and `Recipient` classes.
* Added new file `data.py` that contains various data structures that are used for specific properties.
* Expanded the functionality of the `Recipient` class by adding more properties.
* Added new functions `Named.getNamed` and `Named.getNamedValue` which retrieves a named property or the value of a named property, respectively, based on its name.

**v0.26.3**
* Added new function `MSGFile.save` that causes it and subclasses to raise a `NotImplementedError` if they do not override it.
* Fixed some issues in the changelog.
* Added some additional constants for future use.

**v0.26.2**
* Fixed error in `Message._registerNamedProperty` where I put the exception `KeyError` instead of `AttributeError`.

**v0.26.1**
* Fixed an issue in `openMsg` that would leave the basic `MSGFile` instance open with the function returned, even if the function was returning a specific msg instance.

**v0.26.0**
* [[TheElementalOfDestruction #3](https://github.com/TheElementalOfDestruction/msg-extractor/issues/3)] Implementation of Named properties has finally been added. This allows us to access certain data that was not available to us before through regular methods.
* Added new function `MSGFile.slistDir` that acts like `MSGFile.listDir`, except that it returns a list of strings rather than a list of lists.
* [[TheElementalOfDestruction #6](https://github.com/TheElementalOfDestruction/msg-extractor/issues/6)] Added new function `MSGFile._getTypedStream` which, based on a path formatted in the same way as you would give to `MSGFile._getStringStream`, will return the data in the specified stream without the user needing to know the type before hand. However, if you DO know the type before hand, you can provide this function with one of the values in `constants.FIXED_LENGTH_PROPS_STRING` or `constants.VARIABLE_LENGTH_PROPS_STRING`.
* [[TheElementalOfDestruction #6](https://github.com/TheElementalOfDestruction/msg-extractor/issues/6)] Added new function `MSGFile._getTypedProperty` which, based on a 4 digit hexadecimal string, will return the property in the properties file that matches that string without the type needing to be specified. However, if you DO know the type before hand, you can provide this function with one of the values in `constants.FIXED_LENGTH_PROPS_STRING` or `constants.VARIABLE_LENGTH_PROPS_STRING`.
* [[TheElementalOfDestruction #6](https://github.com/TheElementalOfDestruction/msg-extractor/issues/6)] Added new function `MSGFile._getTypedData` which is a combination of the two previously stated functions.
* Added new function `MSGFile.ExistsTypedProperty` which determines if a property with the specified id exists in the specified location. If you are looking for a property that may be in the properties file of an attachment or a recipient, please use the corresponding function from that class.
* Added an equivalent of the previous 4 functions for the `Recipient` and `Attachment` classes.
* [[TheElementalOfDestruction #2](https://github.com/TheElementalOfDestruction/msg-extractor/issues/2)] Finished partial implementation of `utils.parseType` which was necessary for the proper implementation of named properties. This function is not fully implemented because there are some types we do not fully understand.

**v0.25.3**
* [[TeamMsgExtractor #138](https://github.com/TeamMsgExtractor/msg-extractor/issues/138)] Fixed missing import in `extract_msg/utils.py`.

**v0.25.2**
* [[TeamMsgExtractor #134](https://github.com/TeamMsgExtractor/msg-extractor/issues/134)] Fixed a typo that caused `Message.headerDict` to raise an exception.
* Upgraded code for `Message.headerDict` to avoid accidentally raising a key error if the header is ever missing the "Received" property.
* Fixed an error in the changelog that caused some issue links to link to the wrong place.

**v0.25.1**
* [[TeamMsgExtractor #132](https://github.com/TeamMsgExtractor/msg-extractor/issues/132)] Fixed an issue caused by unfinished code being left in the \_\_main\_\_ file.
* Cleaned up the imports to only be what is needed.

**v0.25.0**
* Added new class `MSGFile`. The `Message` class now inherits from this. This class is the base for all MSG files, not just `Message`s. It somewhat recently came to our attention that MSG files are used for a variety of things, including the storage of contacts, leading us to the next part of the changelog.
* [[TeamMsgExtractor #110](https://github.com/TeamMsgExtractor/msg-extractor/issues/110)] Added new class `Contact` for extracting the data from MSG files storing contacts.
* Added new function `openMsg` to the module to be used to open MSG files in which it is not certain what type of MSG is being opened.
* Modified the `Attachment` class to use the `openMsg` function to open embedded MSG files.
* Added option `delayAttachments` to the `Message` class that will stop it from initializing attachments until the user is ready. This allows users to open `Message`s that have unimplemented attachment types without having to worry about the exception stopping them. This is also an option in the new `openMsg` function.

**v0.24.4**
* Added new property `Message.isRead` to show whether the email has been marked as read.
* Renamed `Message.header_dict` to `Message.headerDict` to better match naming conventions.
* Renamed `Message.message_id` to `Message.messageId` to better match naming conventions.

**v0.24.3**
* Added new close function to the `Message` class to ensure that all embedded `Message` instances get closed as well. Not having this was causing issues with trying to modify the msg file after the user thought that it had been closed.

**v0.24.2**
* Fixed bug that somehow escaped detection that caused certain properties to not work.
* Fixed bug with embedded msg files introduced in v0.24.0

**v0.24.0**
* [[TeamMsgExtractor #107](https://github.com/TeamMsgExtractor/msg-extractor/issues/107)] Rewrote the `Messsage.save` function to fix many errors arising from it and to extend its functionality.
* Added new function `isEmptyString` to check if a string passed to it is `None` or is empty.

**v0.23.4**
* [[TeamMsgExtractor #112](https://github.com/TeamMsgExtractor/msg-extractor/issues/112)] Changed method used to get the message from an exception to make it compatible with Python 2 and 3.
* [[TheElementalOfDestruction #23](https://github.com/TheElementalOfDestruction/msg-extractor/issues/23)] General cleanup and all around improvements of the code.

**v0.23.3**
* Fixed issues in readme.
* [[TheElementalOfDestruction #22](https://github.com/TheElementalOfDestruction/msg-extractor/issues/22)] Updated `dev_classes.Message` to better match the current `Message` class.
* Fixed bad links in changelog.
* [[TeamMsgExtractor #95](https://github.com/TeamMsgExtractor/msg-extractor/issues/95)] Added fallback encoding as well as manual encoding change to `dev_classes.Message`.

**v0.23.1**
* Fixed issue with embedded msg files caused by the changes in v0.23.0.

**v0.23.0**
* [[TeamMsgExtractor #75](https://github.com/TeamMsgExtractor/msg-extractor/issues/75)] & [[TheElementalOfDestruction #19](https://github.com/TheElementalOfDestruction/msg-extractor/issues/19)] Completely rewrote the function `Message._getStringStream`. This was done for two reasons. The first was to make it actually work with msg files that have their strings encoded in a non-Unicode encoding. The second reason was to make it so that it better reflected msg specification which says that ALL strings in a file will be either Unicode or non-Unicode, but not both. Because of the second part, the `prefer` option has been removed.
* As part of fixing the two issues in the previous change, we have added two new properties:
    1. a boolean `Message.areStringsUnicode` which tells if the strings are Unicode encoded.
    2. A string `Message.stringEncoding` which tells what the encoding is. This is used by the `Message._getStringStream` to determine how to decode the data into a string.

**v0.22.1**
* [[TeamMsgExtractor #69](https://github.com/TeamMsgExtractor/msg-extractor/issues/69)] Fixed date format not being up to standard.
* Fixed a minor spelling error in the code.

**v0.22.0**
* [[TheElementalOfDestruction #18](https://github.com/TheElementalOfDestruction/msg-extractor/issues/18)] Added `--validate` option.
* [[TheElementalOfDestruction #16](https://github.com/TheElementalOfDestruction/msg-extractor/issues/16)] Moved all dev code into its own scripts. Use `--dev` to use from the command line.
* [[TeamMsgExtractor #67](https://github.com/TeamMsgExtractor/msg-extractor/issues/67)] Added compatibility module to enforce Unicode os functions.
* Added new function to `Message` class: `Message.sExists`. This function checks if a string stream exists. It's input should be formatted identically to that of `Message._getStringStream`.
* Added new function to `Message` class: `Message.fix_path`. This function will add the proper prefix to the path (if the `prefix` parameter is true) and adjust the path to be a string rather than a list or tuple.
* Added new function to `utils.py`: `get_full_class_name`. This function returns a string containing the module name and the class name of any instance of any class. It is returned in the format of `{module}.{class}`.
* Added a sort of alias of `Message._getStream`, `Message._getStringStream`, `Message.Exists`, and `Message.sExists` to `Attachment` and `Recipient`. These functions run inside the associated attachment directory or recipient directory, respectively.
* Added a fix to an issue introduced in an earlier version caused by accidentally deleting a letter in the code.

**v0.21.0**
* [[TheElementalOfDestruction #12](https://github.com/TheElementalOfDestruction/msg-extractor/issues/12)] Changed debug code to use logging module.
* [[TheElementalOfDestruction #17](https://github.com/TheElementalOfDestruction/msg-extractor/issues/17)] Fixed Attachment class using wrong properties file location in embedded msg files.
* [[TheElementalOfDestruction #11](https://github.com/TheElementalOfDestruction/msg-extractor/issues/11)] Improved handling of command line arguments using argparse module.
* [[TheElementalOfDestruction #16](https://github.com/TheElementalOfDestruction/msg-extractor/issues/16)] Started work on moving developer code into its own script.
* [[TeamMsgExtractor #63](https://github.com/TeamMsgExtractor/msg-extractor/issues/63)] Fixed JSON saving not applying to embedded msg files.
* [[TeamMsgExtractor #55](https://github.com/TeamMsgExtractor/msg-extractor/issues/55)] Added fix for recipient sometimes missing email address.
* [[TeamMsgExtractor #65](https://github.com/TeamMsgExtractor/msg-extractor/issues/65)] Added fix for special characters in recipient names.
* Module now raises a custom exception (instead of just `IOError`) if the input is not a valid OLE file.
* Added `header_dict` property to the `Message` class.
* General minor bug fixes.
* Fixed a section in the `Recipient` class that I have no idea why I did it that way. If errors start randomly occurring with it, this fix is why.

**v0.20.8**
* Fixed a tab issue and parameter type in `message.py`.


**v0.20.7**

* Separated classes into their own files to make things more manageable.
* Placed `__doc__` back inside of `__init__.py`.
* Rewrote the `Prop` class to be two different classes that extend from a base class.
* Made decent progress on completing the `parse_type` function of the `FixedLengthProp` class (formerly a function of the `Prop` class).
* Improved exception handling code throughout most of the module.
* Updated the `.gitignore`.
* Updated README.
* Added `# DEBUG` comments before debugging lines to make them easier to find in the future.
* Added function `create_prop` in `prop.py` which should be used for creating what used to be an instance of the `Prop` class.
* Added more constants to reflect some of the changes made.
* Fixed a major bug that was causing the header to generate after things like "to" and "cc" which would force those fields to not use the header.
* Fixed the debug variable.
* Fixed many small bugs in many of the classes.
* [[TheElementalOfDestruction #13](https://github.com/TheElementalOfDestruction/msg-extractor/issues/13)] Various loose ends to enhance the workflow in the repo.
