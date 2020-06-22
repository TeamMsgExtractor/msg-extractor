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
* [[mattgwwalker # 107](https://github.com/mattgwwalker/msg-extractor/issues/107)] Rewrote the `Messsage.save` function to fix many  errors arising from it and to extend its functionality.
* Added new function `isEmptyString` to check if a string passed to it is `None` or is empty.

**v0.23.4**
* [[mattgwwalker #112](https://github.com/mattgwwalker/msg-extractor/issues/112)] Changed method used to get the message from an exception to make it compatible with Python 2 and 3.
* [[TheElementalOfCreation #23](https://github.com/TheElementalOfCreation/msg-extractor/issues/23)] General cleanup and all around improvements of the code.

**v0.23.3**
* Fixed issues in readme.
* [[TheElementalOfCreation #22](https://github.com/TheElementalOfCreation/msg-extractor/issues/22)] Updated `dev_classes.Message` to better match the current `Message` class.
* Fixed bad links in changelog.
* [[mattgwwalker #95](https://github.com/mattgwwalker/msg-extractor/issues/95)] Added fallback encoding as well as manual encoding change to `dev_classes.Message`.

**v0.23.1**
* Fixed issue with embedded msg files caused by the changes in v0.23.0.

**v0.23.0**
* [[mattgwwalker #75](https://github.com/mattgwwalker/msg-extractor/issues/75)] & [[TheElementalOfCreation #19](https://github.com/TheElementalOfCreation/msg-extractor/issues/19)] Completely rewrote the function `Message._getStringStream`. This was done for two reasons. The first was to make it actually work with msg files that have their strings encoded in a non-Unicode encoding. The second reason was to make it so that it better reflected msg specification which says that ALL strings in a file will be either Unicode or non-Unicode, but not both. Because of the second part, the `prefer` option has been removed.
* As part of fixing the two issues in the previous change, we have added two new properties:
    1. a boolean `Message.areStringsUnicode` which tells if the strings are unicode encoded.
    2. A string `Message.stringEncoding` which tells what the encoding is. This is used by the `Message._getStringStream` to determine how to decode the data into a string.

**v0.22.1**
* [[mattgwwalker #69](https://github.com/mattgwwalker/msg-extractor/issues/69)] Fixed date format not being up to standard.
* Fixed a minor spelling error in the code.

**v0.22.0**
* [[TheElementalOfCreation #18](https://github.com/TheElementalOfCreation/msg-extractor/issues/18)] Added `--validate` option.
* [[TheElementalOfCreation #16](https://github.com/TheElementalOfCreation/msg-extractor/issues/16)] Moved all dev code into its own scripts. Use `--dev` to use from the command line.
* [[mattgwwalker #67](https://github.com/mattgwwalker/msg-extractor/issues/67)] Added compatability module to enforce unicode os functions.
* Added new function to `Message` class: `Message.sExists`. This function checks if a string stream exists. It's input should be formatted identically to that of `Message._getStringSteam`.
* Added new function to `Message` class: `Message.fix_path`. This function will add the proper prefix to the path (if the `prefix` parameter is true) and adjust the path to be a string rather than a list or tuple.
* Added new function to `utils.py`: `get_full_class_name`. This function returns a string containing the module name and the class name of any instance of any class. It is returned in the format of `{module}.{class}`.
* Added a sort of alias of `Message._getStream`, `Message._getStringStream`, `Message.Exists`, and `Message.sExists` to `Attachment` and `Recipient`. These functions run inside the associated attachment directory or recipient directory, respectively.
* Added a fix to an issue introduced in an earlier version caused by accidentally deleting a letter in the code.

**v0.21.0**
* [[TheElementalOfCreation #12](https://github.com/TheElementalOfCreation/msg-extractor/issues/12)] Changed debug code to use logging module.
* [[TheElementalOfCreation #17](https://github.com/TheElementalOfCreation/msg-extractor/issues/17)] Fixed Attachment class using wrong properties file location in embedded msg files.
* [[TheElementalOfCreation #11](https://github.com/TheElementalOfCreation/msg-extractor/issues/11)] Improved handling of command line arguments using argparse module.
* [[TheElementalOfCreation #16](https://github.com/TheElementalOfCreation/msg-extractor/issues/16)] Started work on moving developer code into its own script.
* [[mattgwwalker #63](https://github.com/mattgwwalker/msg-extractor/issues/63)] Fixwed json saving not applying to embedded msg files.
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
* [[TheElementalOfCreation #13](https://github.com/TheElementalOfCreation/msg-extractor/issues/13)] Various loose ends to enhance the workflow in the repo.
