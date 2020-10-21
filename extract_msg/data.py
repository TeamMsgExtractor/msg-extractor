"""
Various small data structures used in msg extractor.
"""

from extract_msg import constants


class PermanentEntryID(object):
    def __init__(self, data):
        super(PermanentEntryID, self).__init__()
        self.__data = data
        unpacked = constants.STPEID.unpack(data[:28])
        if unpacked[0] != 0:
            raise TypeError('Not a PermanentEntryID (expected 0, got {}).'.format(unpacked[0]))
        self.__providerUID = unpacked[1]
        self.__displayTypeString = unpacked[2]
        self.__distinguishedName = data[28:-1].decode('ascii') # Cut off the null character at the end and decode the data as ascii

    @property
    def data(self):
        """
        Returns the raw data used to generate this instance.
        """
        return self.__data

    @property
    def displayTypeString(self):
        """
        Returns the display type string. This will be one of the display type constants.
        """
        return self.__displayTypeString

    @property
    def distinguishedName(self):
        """
        Returns the distinguished name.
        """
        return self.__distinguishedName

    @property
    def providerUID(self):
        """
        Returns the provider UID.
        """
        return self.__providerUID
