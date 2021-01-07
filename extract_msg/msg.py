import codecs
import copy
import logging

import olefile

from extract_msg import constants
from extract_msg.attachment import Attachment
from extract_msg.named import Named
from extract_msg.prop import FixedLengthProp, VariableLengthProp
from extract_msg.properties import Properties
from extract_msg.utils import divide, getEncodingName, has_len, inputToMsgpath, inputToString, msgpathToString, parseType, properHex, verifyPropertyId, verifyType, windowsUnicode
from extract_msg.exceptions import InvalidFileFormatError, MissingEncodingError



logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

class MSGFile(olefile.OleFileIO):
    """
    Parser for .msg files
    """
    def __init__(self, path, prefix = '', attachmentClass = Attachment, filename = None, overrideEncoding = None, attachmentErrorBehavior = constants.ATTACHMENT_ERROR_THROW):
        """
        :param path: path to the msg file in the system or is the raw msg file.
        :param prefix: used for extracting embeded msg files
            inside the main one. Do not set manually unless
            you know what you are doing.
        :param attachmentClass: optional, the class the MSGFile object
            will use for attachments. You probably should
            not change this value unless you know what you
            are doing.
        :param filename: optional, the filename to be used by default when saving.
        :param overrideEncoding: optional, an encoding to use instead of the one
            specified by the msg file. Do not report encoding errors caused by this.
        """
        # WARNING DO NOT MANUALLY MODIFY PREFIX. Let the program set it.
        self.__path = path
        self.__attachmentClass = attachmentClass
        if not (constants.ATTACHMENT_ERROR_THROW <= attachmentErrorBehavior <= constants.ATTACHMENT_ERROR_BROKEN):
            raise ValueError("`attachmentErrorBehavior` must be ATTACHMENT_ERROR_THROW, ATTACHMENT_ERROR_NOT_IMPLEMENTED, or ATTACHMENT_ERROR_BROKEN.")
        self.__attachmentErrorBehavior = attachmentErrorBehavior
        if overrideEncoding is not None:
            codecs.lookup(overrideEncoding)
            logger.warning('You have chosen to override the string encoding. Do not report encoding errors caused by this.')
            self.__stringEncoding = overrideEncoding
        self.__overrideEncoding = overrideEncoding

        try:
            olefile.OleFileIO.__init__(self, path)
        except IOError as e:    # py2 and py3 compatible
            logger.error(e)
            if str(e) == 'not an OLE2 structured storage file':
                raise InvalidFileFormatError(e)
            else:
                raise

        prefixl = []
        tmp_condition = prefix != ''
        if tmp_condition:
            try:
                prefix = inputToString(prefix, 'utf-8')
            except:
                try:
                    prefix = '/'.join(prefix)
                except:
                    raise TypeError('Invalid prefix type: ' + str(type(prefix)) +
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
            filename = self._getStringStream(prefixl[:-1] + ['__substg1.0_3001'], prefix=False)
        if filename is not None:
            self.filename = filename
        elif has_len(path):
            if len(path) < 1536:
                self.filename = path
            else:
                self.filename = None
        else:
            self.filename = None

    def _ensureSet(self, variable, streamID, stringStream = True):
        """
        Ensures that the variable exists, otherwise will set it using the specified stream.
        After that, return said variable.
        If the specified stream is not a string stream, make sure to set :param string stream: to False.
        """
        try:
            return getattr(self, variable)
        except AttributeError:
            if stringStream:
                value = self._getStringStream(streamID)
            else:
                value = self._getStream(streamID)
            setattr(self, variable, value)
            return value

    def _ensureSetNamed(self, variable, propertyName):
        """
        Ensures that the variable exists, otherwise will set it using the named property.
        After that, return said variable.
        """
        try:
            return getattr(self, variable)
        except AttributeError:
            value = self.named.getNamedValue(propertyName)
            setattr(self, variable, value)
            return value

    def _ensureSetProperty(self, variable, propertyName):
        """
        Ensures that the variable exists, otherwise will set it using the property.
        After that, return said variable.
        """
        try:
            return getattr(self, variable)
        except AttributeError:
            try:
                value = self.mainProperties[propertyName].value
            except (KeyError, AttributeError):
                value = None
            setattr(self, variable, value)
            return value

    def _getStream(self, filename, prefix = True):
        """
        Gets a binary representation of the requested filename.

        This should ALWAYS return a bytes object (string in python 2)
        """
        filename = self.fix_path(filename, prefix)
        if self.exists(filename):
            with self.openstream(filename) as stream:
                return stream.read()
        else:
            logger.info('Stream "{}" was requested but could not be found. Returning `None`.'.format(filename))
            return None

    def _getStringStream(self, filename, prefix = True):
        """
        Gets a string representation of the requested filename.

        Rather than the full filename, you should only feed this
        function the filename sans the type. So if the full name
        is "__substg1.0_001A001F", the filename this function
        should receive should be "__substg1.0_001A".

        This should ALWAYS return a string (Unicode in python 2)
        """

        filename = self.fix_path(filename, prefix)
        if self.areStringsUnicode:
            return windowsUnicode(self._getStream(filename + '001F', prefix = False))
        else:
            tmp = self._getStream(filename + '001E', prefix = False)
            return None if tmp is None else tmp.decode(self.stringEncoding)

    def _getTypedData(self, id, _type = None, prefix = True):
        """
        Gets the data for the specified id as the type that it is
        supposed to be. :param id: MUST be a 4 digit hexadecimal
        string.

        If you know for sure what type the data is before hand,
        you can specify it as being one of the strings in the
        constant FIXED_LENGTH_PROPS_STRING or
        VARIABLE_LENGTH_PROPS_STRING.
        """
        verifyPropertyId(id)
        id = id.upper()
        found, result = self._getTypedStream('__substg1.0_' + id, prefix, _type)
        if found:
            return result
        else:
            found, result = self._getTypedProperty(id, _type)
            return result if found else None

    def _getTypedProperty(self, propertyID, _type = None):
        """
        Gets the property with the specified id as the type that it
        is supposed to be. :param id: MUST be a 4 digit hexadecimal
        string.

        If you know for sure what type the property is before hand,
        you can specify it as being one of the strings in the
        constant FIXED_LENGTH_PROPS_STRING or
        VARIABLE_LENGTH_PROPS_STRING.
        """
        verifyPropertyId(propertyID)
        verifyType(_type)
        propertyID = propertyID.upper()
        for x in (propertyID + _type,) if _type is not None else self.mainProperties:
            if x.startswith(propertyID):
                prop = self.mainProperties[x]
                return True, (prop.value if isinstance(prop, FixedLengthProp) else prop)
        return False, None

    def _getTypedStream(self, filename, prefix = True, _type = None):
        """
        Gets the contents of the specified stream as the type that
        it is supposed to be.

        Rather than the full filename, you should only feed this
        function the filename sans the type. So if the full name
        is "__substg1.0_001A001F", the filename this function
        should receive should be "__substg1.0_001A".

        If you know for sure what type the stream is before hand,
        you can specify it as being one of the strings in the
        constant FIXED_LENGTH_PROPS_STRING or
        VARIABLE_LENGTH_PROPS_STRING.

        If you have not specified the type, the type this function
        returns in many cases cannot be predicted. As such, when
        using this function it is best for you to check the type
        that it returns. If the function returns None, that means
        it could not find the stream specified.
        """
        verifyType(_type)
        filename = self.fix_path(filename, prefix)
        for x in (filename + _type,) if _type is not None else self.slistDir():
            if x.startswith(filename) and x.find('-') == -1:
                contents = self._getStream(x, False)
                if len(contents) == 0:
                    return True, None # We found the file, but it was empty.
                extras = []
                _type = x[-4:]
                if x[-4] == '1': # It's a multiple
                    if _type in ('101F', '101E'):
                        streams = len(contents) // 4 # These lengths are normal.
                    elif _type == '1102':
                        streams = len(contents) // 8 # These lengths have 4 0x00 bytes at the end for seemingly no reason. They are "reserved" bytes
                    elif _type in ('1002', '1003', '1004', '1005', '1007', '1014', '1040', '1048'):
                        try:
                            streams = self.mainProperties[x[-8:]].realLength
                        except:
                            logger.error('Could not find matching VariableLengthProp for stream {}'.format(x))
                            streams = len(contents) // (2 if _type in constants.MULTIPLE_2_BYTES else 4 if _type in constants.MULTIPLE_4_BYTES else 8 if _type in constants.MULTIPLE_8_BYTES else 16)
                    else:
                        raise NotImplementedError('The stream specified is of type {}. We don\'t currently understand exactly how this type works. If it is mandatory that you have the contents of this stream, please create an issue labled "NotImplementedError: _getTypedStream {}".'.format(_type, _type))
                    if _type in ('101F', '101E', '1102'):
                        if self.Exists(x + '-00000000', False):
                            for y in range(streams):
                                if self.Exists(x + '-' + properHex(y, 8), False):
                                    extras.append(self._getStream(x + '-' + properHex(y, 8), False))
                    elif _type in ('1002', '1003', '1004', '1005', '1007', '1014', '1040', '1048'):
                        extras = divide(contents, (2 if _type in constants.MULTIPLE_2_BYTES else 4 if _type in constants.MULTIPLE_4_BYTES else 8 if _type in constants.MULTIPLE_8_BYTES else 16))
                        contents = streams
                return True, parseType(int(_type, 16), contents, self.stringEncoding, extras)
        return False, None # We didn't find the stream.

    def _registerNamedProperty(self, entry, _type, name = None):
        """
        FOR INTERNAL USE ONLY! DO NOT CALL MANUALLY!

        Function to allow things like attachments in subclasses to have their own named properties.
        """
        pass

    def debug(self):
        for dir_ in self.listDir():
            if dir_[-1].endswith('001E') or dir_[-1].endswith('001F'):
                print('Directory: ' + str(dir_[:-1]))
                print('Contents: {}'.format(self._getStream(dir_)))

    def Exists(self, inp, prefix = True):
        """
        Checks if :param inp: exists in the msg file.
        """
        inp = self.fix_path(inp, prefix)
        return self.exists(inp)

    def sExists(self, inp, prefix = True):
        """
        Checks if string stream :param inp: exists in the msg file.
        """
        inp = self.fix_path(inp, prefix)
        return self.exists(inp + '001F') or self.exists(inp + '001E')

    def ExistsTypedProperty(self, id, location = None, _type = None, prefix = True, propertiesInstance = None):
        """
        Determines if the stream with the provided id exists in the location specified.
        If no location is specified, the root directory is searched. The return of this
        function is 2 values, the first being a boolean for if anything was found, and
        the second being how many were found.

        Because of how this function works, any folder that contains it's own
        "__properties_version1.0" file should have this function called from it's class.
        """
        verifyPropertyId(id)
        verifyType(_type)
        id = id.upper()
        if propertiesInstance is None:
            propertiesInstance = self.mainProperties
        prefixList = self.prefixList if prefix else []
        if location is not None:
            prefixList.append(location)
        prefixList = inputToMsgpath(prefixList)
        usableid = id + _type if _type is not None else id
        found_number = 0
        found_streams = []
        for item in self.listDir():
            if len(item) > len(prefixList):
                if item[len(prefixList)].startswith('__substg1.0_' + usableid) and item[len(prefixList)] not in found_streams:
                    found_number += 1
                    found_streams.append(item[len(prefixList)])
        for x in propertiesInstance:
            if x.startswith(usableid):
                already_found = False
                for y in found_streams:
                    if y.endswith(x):
                        already_found = True
                        break
                if not already_found:
                    found_number += 1
        return (found_number > 0), found_number

    def fix_path(self, inp, prefix = True):
        """
        Changes paths so that they have the proper
        prefix (should :param prefix: be True) and
        are strings rather than lists or tuples.
        """
        inp = msgpathToString(inp)
        if prefix:
            inp = self.__prefix + inp
        return inp

    def listDir(self, streams = True, storages = False):
        """
        Replacement for OleFileIO.listdir that runs at the current prefix directory.
        """
        temp = self.listdir(streams, storages)
        if self.__prefix == '':
            return temp
        prefix = self.__prefix.split('/')
        if prefix[-1] == '':
            prefix.pop()
        out = []
        for x in temp:
            good = True
            if len(x) <= len(prefix):
                good = False
            if good:
                for y in range(len(prefix)):
                    if x[y] != prefix[y]:
                        good = False
            if good:
                out.append(x)
        return out

    def slistDir(self, streams = True, storages = False):
        """
        Replacement for OleFileIO.listdir that runs at the current prefix directory.
        Returns a list of strings instead of lists.
        """
        return [msgpathToString(x) for x in self.listDir(streams, storages)]

    def save(self, *args, **kwargs):
        raise NotImplementedError('Saving is not yet supported for the {} class'.format(self.__class__.__name__))

    @property
    def areStringsUnicode(self):
        """
        Returns a boolean telling if the strings are unicode encoded.
        """
        try:
            return self.__bStringsUnicode
        except AttributeError:
            if self.mainProperties.has_key('340D0003'):
                if (self.mainProperties['340D0003'].value & 0x40000) != 0:
                    self.__bStringsUnicode = True
                    return self.__bStringsUnicode
            self.__bStringsUnicode = False
            return self.__bStringsUnicode

    @property
    def attachmentClass(self):
        """
        Returns the Attachment class being used, should you need to use it externally for whatever reason.
        """
        return self.__attachmentClass

    @property
    def attachmentErrorBehavior(self):
        """
        The behavior to follow when an attachment raises an exception. Will be one
        of the following values:
        ATTACHMENT_ERROR_THROW: Don't catch exceptions.
        ATTACHMENT_ERROR_NOT_IMPLEMENTED: Catch NotImplementedError exceptions.
        ATTACHMENT_ERROR_BROKEN: Catch all exceptions.
        """
        return self.__attachmentErrorBehavior

    @property
    def classType(self):
        """
        The class type of the MSG file.
        """
        return self._ensureSet('_classType', '__substg1.0_001A')

    @property
    def importance(self):
        """
        The specified importance of the msg file.
        """
        return self._ensureSetProperty('_importance', '00170003')

    @property
    def mainProperties(self):
        """
        Returns the Properties instance used by the MSGFile instance.
        """
        try:
            return self._prop
        except AttributeError:
            self._prop = Properties(self._getStream('__properties_version1.0'),
                                    constants.TYPE_MESSAGE if self.prefix == '' else constants.TYPE_MESSAGE_EMBED)
            return self._prop

    @property
    def named(self):
        """
        The main named properties instance for this file.
        """
        try:
            return self.__namedProperties
        except AttributeError:
            self.__namedProperties = Named(self)
            return self.__namedProperties

    @property
    def overrideEncoding(self):
        """
        Returns None is the encoding has not been overriden, otherwise returns
        the encoding.
        """
        return self.__overrideEncoding

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
    def priority(self):
        """
        The specified priority of the msg file.
        """
        return self._ensureSetProperty('_priority', '00260003')

    @property
    def sensitivity(self):
        """
        The specified sensitivity of the msg file.
        """
        return self._ensureSetProperty('_sensitivity', '00360003')

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
                if not self.mainProperties.has_key('3FFD0003'):
                    raise MissingEncodingError('Encoding property not found')
                enc = self.mainProperties['3FFD0003'].value
                # Now we just need to translate that value
                self.__stringEncoding = getEncodingName(enc)
                return self.__stringEncoding
