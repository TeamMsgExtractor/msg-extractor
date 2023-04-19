import unittest


class ExceptionExpectedTestCase(unittest.TestCase):
    def assertException(self, callable, excType, message = None, subclassesAllowed = False):
        """
        Assert that the specific exception is raised.

        Returns a wrapper around the function that can be called with the
        original function arguments. The wrapper will catch all normal
        exceptions and check the type against the provided one. If a message is
        specified, confirms that the message matches exactly.

        if :param subclassesAllowed: is True, any subclasses of the exception
        type will be accepted. Otherwise a subclass is considered a failure.
        """
        def wrapper(*args, **kwargs):
            try:
                callable(*args, **kwargs)
            except Exception as e:
                if isinstance(e, excType):
                    if subclassesAllowed or not issubclass(e, excType):
                        return
                # If we get here, raise the exception to have the function fail.
                raise

        return wrapper
