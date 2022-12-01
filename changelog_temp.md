Temporary location for the changelog entry to ensure it doesn't conflict.

**v0.??.??**
* Added new submodule `custom_attachments`. This submodule provides an extendable way to handle custom attachment types, attachment types whose structure and formatting are not defined in the Microsoft documentation for MSG files.
* Added new property `AttachmentBase.clsid` which returns the listed CLSID value of the data stream/storage of the attachment.
* Changed internal behavior of `MSGFile.attachments`. This should not cause any noticeable changes to the output.
