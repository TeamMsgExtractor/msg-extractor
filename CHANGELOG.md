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
* [[mattgwwalker #138](https://github.com/mattgwwalker/msg-extractor/issues/138)] Fixed missing import in extract_msg/utils.py.

**v0.25.2**
* [[mattgwwalker #134](https://github.com/mattgwwalker/msg-extractor/issues/134)] Fixed a typo that caused `Message.headerDict` to raise an exception.
* Upgraded code for `Message.headerDict` to avoid accidentally raising a key error if the header is ever missing the "Received" property.
* Fixed an error in the changelog that caused some issue links to link to the wrong place.

**v0.25.1**
* [[mattgwwalker #132](https://github.com/mattgwwalker/msg-extractor/issues/132)] Fixed an issue caused by unfinished code being left in the \_\_main\_\_ file.
* Cleaned up the imports to only be what is needed.

**v0.25.0**
* Added new class `MSGFile`. The `Message` class now inherits from this. This class is the base for all MSG files, not just `Message`s. It somewhat recently came to our attention that MSG files are used for a variety of things, including the storage of contacts, leading us to the next part of the changelog.
* [[mattgwwalker #110](https://github.com/mattgwwalker/msg-extractor/issues/110)] Added new class `Contact` for extracting the data from MSG files storing contacts.
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
* [[mattgwwalker #107](https://github.com/mattgwwalker/msg-extractor/issues/107)] Rewrote the `Messsage.save` function to fix many errors arising from it and to extend its functionality.
* Added new function `isEmptyString` to check if a string passed to it is `None` or is empty.

**v0.23.4**
* [[mattgwwalker #112](https://github.com/mattgwwalker/msg-extractor/issues/112)] Changed method used to get the message from an exception to make it compatible with Python 2 and 3.
* [[TheElementalOfDestruction #23](https://github.com/TheElementalOfDestruction/msg-extractor/issues/23)] General cleanup and all around improvements of the code.

**v0.23.3**
* Fixed issues in readme.
* [[TheElementalOfDestruction #22](https://github.com/TheElementalOfDestruction/msg-extractor/issues/22)] Updated `dev_classes.Message` to better match the current `Message` class.
* Fixed bad links in changelog.
* [[mattgwwalker #95](https://github.com/mattgwwalker/msg-extractor/issues/95)] Added fallback encoding as well as manual encoding change to `dev_classes.Message`.

**v0.23.1**
* Fixed issue with embedded msg files caused by the changes in v0.23.0.

**v0.23.0**
* [[mattgwwalker #75](https://github.com/mattgwwalker/msg-extractor/issues/75)] & [[TheElementalOfDestruction #19](https://github.com/TheElementalOfDestruction/msg-extractor/issues/19)] Completely rewrote the function `Message._getStringStream`. This was done for two reasons. The first was to make it actually work with msg files that have their strings encoded in a non-Unicode encoding. The second reason was to make it so that it better reflected msg specification which says that ALL strings in a file will be either Unicode or non-Unicode, but not both. Because of the second part, the `prefer` option has been removed.
* As part of fixing the two issues in the previous change, we have added two new properties:
    1. a boolean `Message.areStringsUnicode` which tells if the strings are Unicode encoded.
    2. A string `Message.stringEncoding` which tells what the encoding is. This is used by the `Message._getStringStream` to determine how to decode the data into a string.

**v0.22.1**
* [[mattgwwalker #69](https://github.com/mattgwwalker/msg-extractor/issues/69)] Fixed date format not being up to standard.
* Fixed a minor spelling error in the code.

**v0.22.0**
* [[TheElementalOfDestruction #18](https://github.com/TheElementalOfDestruction/msg-extractor/issues/18)] Added `--validate` option.
* [[TheElementalOfDestruction #16](https://github.com/TheElementalOfDestruction/msg-extractor/issues/16)] Moved all dev code into its own scripts. Use `--dev` to use from the command line.
* [[mattgwwalker #67](https://github.com/mattgwwalker/msg-extractor/issues/67)] Added compatibility module to enforce Unicode os functions.
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
* [[mattgwwalker #63](https://github.com/mattgwwalker/msg-extractor/issues/63)] Fixed JSON saving not applying to embedded msg files.
* [[mattgwwalker #55](https://github.com/mattgwwalker/msg-extractor/issues/55)] Added fix for recipient sometimes missing email address.
* [[mattgwwalker #65](https://github.com/mattgwwalker/msg-extractor/issues/65)] Added fix for special characters in recipient names.
* Module now raises a custom exception (instead of just `IOError`) if the input is not a valid OLE file.
* Added `header_dict` property to the `Message` class.
* General minor bug fixes.
* Fixed a section in the `Recipient` class that I have no idea why I did it that way. If errors start randomly occurring with it, this fix is why.

**v0.20.8**
* Fixed a tab issue and parameter type in `messages.py`.


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
