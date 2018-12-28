**v0.21.0**
* [[Syncurity #7](https://github.com/Syncurity/msg-extractor/issues/7)] Changed debug code to use logging module.
* [[Syncurity #26](https://github.com/Syncurity/msg-extractor/issues/26)] Fixed Attachment class using wrong properties file location in embedded msg files.
* [[Syncurity #4](https://github.com/Syncurity/msg-extractor/issues/4)] Improved handling of command line arguments using argparse module.
* [[Syncurity #24](https://github.com/Syncurity/msg-extractor/issues/24)] Started work on moving developer code into its own script.
* [[mattgwwalker #63](https://github.com/mattgwwalker/msg-extractor/issues/63)] Fixwed json saving not applying to embedded msg files.
* [[mattgwwalker #55](https://github.com/mattgwwalker/msg-extractor/issues/55)] Added fix for recipient sometimes missing email address.
* [[mattgwwalker #65](https://github.com/mattgwwalker/msg-extractor/issues/65)] Added fix for special characters in recipient names.
* Module now raises a custom exception (instead of just `IOError`) if the input is not a valid OLE file.
* Added `header_dict` property to the `Message` class.
* General minor bug fixes.
* Fixed a section in the `Recipient` class that I have no idea why I did it that way. If errors start randomly occurring with it, this fix is why.

**v0.20.8**
* Fixed a tab issue and parameter type in `messages.py`.


**v0.20.7:**

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
* [[Syncurity #11](https://github.com/Syncurity/msg-extractor/issues/11)] Various loose ends to enhance the workflow in the repo.
