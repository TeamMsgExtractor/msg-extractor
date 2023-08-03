**v0.?.?**
* Changed imports in `message_base.py` to help with type checkers.
* Added new function `MessageBase.asEmailMessage` which will convert the `MessageBase` instance, if possible, to an `email.message.EmailMessage` object. If an embedded MSG file on a `MessageBase` object is of a class that does not have this function, it will simply be attached to the message as bytes.
* Changed from using `email.parser.EmailParser` to `email.parser.HeaderParser` in `MessageBase.header`.
* Changed some of the internal code for `MessageBase.header`. This should improve usage of it, and should not have any notiocable negative changes. You man notice some of the values parse slightly differently, but this effect should be mostly supressed.