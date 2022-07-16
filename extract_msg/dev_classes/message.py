import copy
import logging
import olefile

from ..enums import PropertiesType
from ..dev_classes.attachment import Attachment
from ..properties import Properties
from ..recipient import Recipient
from ..utils import hasLen, inputToString, windowsUnicode


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

class Message(olefile.OleFileIO):
    """
    Developer version of the `extract_msg.message.Message` class.
    Useful for malformed msg files.
    """

    def __init__(self, path, prefix = '', filename = None):
        """
        :param path: path to the msg file in the system or is the raw msg file.
        :param prefix: used for extracting embedded msg files
            inside the main one. Do not set manually unless
            you know what you are doing.
        """
        logger.log(5, f'Prefix: {prefix}')
        self.__path = path
        super().__init__(path)
        prefixl = []
        tmp_condition = prefix != ''
        if tmp_condition:
            prefix = inputToString(prefix, 'utf-8')
            try:
                prefix = '/'.join(prefix)
            except TypeError:
                raise TypeError(f'Invalid prefix type: {type(prefix)}' +
                                '\n(This was probably caused by you setting it manually).')
            prefix = prefix.replace('\\', '/')
            g = prefix.split('/')
            if g[-1] == '':
                g.pop()
            prefixl = g
            if prefix[-1] != '/':
                prefix += '/'
        self.__prefix = prefix
        self.__prefixList = prefixl

        if tmp_condition:
            filename = self._getStringStream(prefixl[:-1] + ['__substg1.0_3001'], prefix = False)
        if filename is not None:
            self.filename = filename
        else:
            logger.log(5, f':param path: has __len__ attribute?: {hasLen(path)}')
            if hasLen(path):
                if len(path) < 1536:
                    self.filename = path
                    logger.log(5, f':param path: length is {len(path)}; Using :param path: as file path')
                else:
                    logger.log(5, f':param path: length is {len(path)}; Using :param path: as raw msg stream')
                    self.filename = None
            else:
                self.filename = None

        self.props
        recipientDirs = []

        for dir_ in self.listDir():
            if dir_[len(self.__prefixList)].startswith('__recip') and \
                    dir_[len(self.__prefixList)] not in recipientDirs:
                recipientDirs.append(dir_[len(self.__prefixList)])

        self.recipients
        self.attachments
        self.date

    def _getStream(self, filename, prefix = True):
        filename = self.fix_path(filename, prefix)
        if self.exists(filename):
            stream = self.openstream(filename)
            return stream.read()
        else:
            logger.info(f'Stream "{filename}" was requested but could not be found. Returning `None`.')
            return None

    def _getStringStream(self, filename, prefix = True):
        """
        Gets a string representation of the requested filename.
        This should ALWAYS return a string.
        """

        filename = self.fixPath(filename, prefix)
        if self.areStringsUnicode:
            return windowsUnicode(self._getStream(filename + '001F', prefix = False))
        else:
            tmp = self._getStream(filename + '001E', prefix = False)
            return None if tmp is None else tmp.decode(self.stringEncoding)

    def exists(self, filename):
        """
        Checks if :param filename: exists in the msg file.
        """
        filename = self.fixPath(filename)
        return self.exists(filename)

    def sExists(self, filename):
        """
        Checks if string stream :param filename: exists in the msg file.
        """
        filename = self.fixPath(filename)
        return self.exists(filename + '001F') or self.exists(filename + '001E')

    def fixPath(self, filename, prefix=True):
        """
        Changes paths so that they have the proper
        prefix (should :param prefix: be True) and
        are strings rather than lists or tuples.
        """
        if isinstance(filename, (list, tuple)):
            filename = '/'.join(filename)
        if prefix:
            filename = self.__prefix + filename
        return filename

    def listDir(self, streams = True, storages = False):
        """
        Replacement for OleFileIO.listdir that runs at the current prefix directory.
        """
        entries = self.listdir(streams, storages)
        # If we are in the top level MSG already, we can just return.
        if not self.__prefix:
            return entries

        # Get a list for the prefix.
        prefix = self.__prefix.split('/')
        if prefix[-1] == '':
            prefix.pop()

        out = []
        for pathEntry in entries:
            good = True
            # If the entry we are looking at is not longer then the prefix, it's not good.
            if len(pathEntry) <= len(prefix):
                continue

            for index, entry in enumerate(prefix):
                if pathEntry[index] != entry:
                    good = False

            if good:
                out.append(pathEntry)

        return out

    @property
    def areStringsUnicode(self):
        """
        Returns a boolean telling if the strings are unicode encoded.
        """
        try:
            return self.__bStringsUnicode
        except AttributeError:
            if self.props.has_key('340D0003'):
                if (self.props['340D0003'].value & 0x40000) != 0:
                    self.__bStringsUnicode = True
                    return self.__bStringsUnicode
            self.__bStringsUnicode = False
            return self.__bStringsUnicode

    @property
    def attachments(self):
        """
        Returns a list of all attachments.
        """
        try:
            return self._attachments
        except AttributeError:
            # Get the attachments
            attachmentDirs = []

            for dir_ in self.listDir():
                if dir_[len(self.__prefixList)].startswith('__attach') and \
                        dir_[len(self.__prefixList)] not in attachmentDirs:
                    attachmentDirs.append(dir_[len(self.__prefixList)])

            self._attachments = []

            for attachmentDir in attachmentDirs:
                self._attachments.append(Attachment(self, attachmentDir))

            return self._attachments

    @property
    def date(self):
        """
        Returns the send date, if it exists.
        """
        try:
            return self._date
        except AttributeError:
            self._date = self._prop.date
            return self._date

    @property
    def props(self):
        """
        Returns the Properties instance used by the Message instance.
        """
        try:
            return self._prop
        except AttributeError:
            self._prop = Properties(self._getStream('__properties_version1.0'),
                                    PropertiesType.MESSAGE if self.__prefix == '' else PropertiesType.MESSAGE_EMBED)
            return self._prop

    @property
    def path(self):
        """
        Returns the message path if generated from a file,
        otherwise returns the data used to generate the
        Message instance.
        """
        return self.__path

    @property
    def prefix(self):
        """
        Returns the prefix of the Message instance.
        Intended for developer use.
        """
        return self.__prefix

    @property
    def prefixList(self):
        """
        Returns the prefix list of the Message instance.
        Intended for developer use.
        """
        return copy.deepcopy(self.__prefixList)

    @property
    def recipients(self):
        """
        Returns a list of all recipients.
        """
        try:
            return self._recipients
        except AttributeError:
            # Get the recipients
            recipientDirs = []

            for dir_ in self.listDir():
                if dir_[len(self.__prefixList)].startswith('__recip') and \
                        dir_[len(self.__prefixList)] not in recipientDirs:
                    recipientDirs.append(dir_[len(self.__prefixList)])

            self._recipients = []

            for recipientDir in recipientDirs:
                self._recipients.append(Recipient(recipientDir, self))

            return self._recipients

    @property
    def stringEncoding(self):
        try:
            return self.__stringEncoding
        except AttributeError:
            # We need to calculate the encoding
            # Let's first check if the encoding will be unicode:
            if self.areStringsUnicode:
                self.__stringEncoding = "utf-16-le"
                return self.__stringEncoding
            else:
                # Well, it's not unicode. Now we have to figure out what it IS.
                if not self.props.has_key('3FFD0003'):
                    logger.error("String encoding is not unicode, but was also not specified. Malformed MSG file detected. Defaulting to utf-8")
                    self.__stringEncoding = 'utf-8'
                    return self.__stringEncoding
                enc = self.props['3FFD0003'].value
                # Now we just need to translate that value
                # Now, this next line SHOULD work, but it is possible that it might not...
                self.__stringEncoding = str(enc)
                return self.__stringEncoding

    @stringEncoding.setter
    def stringEncoding(self, enc):
        self.__stringEncoding = enc
