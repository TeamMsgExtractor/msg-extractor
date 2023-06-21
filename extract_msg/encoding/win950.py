"""
Support for Microsoft's implementation of CP950 (core python has bad support
for it).
"""

__all__ = [
    'getregentry',
]


# We use a similar format to what I've seen in core Python encoding files.
import codecs

from .utils import variableByteDecode, variableByteEncode
from ._win950_dec import decodingTable

### Codec APIs

class Codec(codecs.Codec):
    def encode(self, text, errors='strict'):
        return variableByteEncode('windows-950', text, errors, encodingTable)

    def decode(self, data, errors='strict'):
        return variableByteDecode('windows-950', data, errors, decodingTable)

class IncrementalEncoder(codecs.IncrementalEncoder):
    def encode(self, text, final=False):
        return variableByteEncode('windows-950', text, self.errors, encodingTable)[0]

class IncrementalDecoder(codecs.IncrementalDecoder):
    def decode(self, data, final=False):
        return variableByteDecode('windows-950', data, self.errors, decodingTable)[0]

class StreamWriter(Codec, codecs.StreamWriter):
    pass

class StreamReader(Codec, codecs.StreamReader):
    pass

### encodings module API

def getregentry():
    return codecs.CodecInfo(
        name='windows-950',
        encode=Codec().encode,
        decode=Codec().decode,
        incrementalencoder=IncrementalEncoder,
        incrementaldecoder=IncrementalDecoder,
        streamwriter=StreamWriter,
        streamreader=StreamReader,
    )


### Encoding table
encodingTable = {value : bytes((key,)) if key < 256
                 else bytes((key >> 8, key & 0xFF))
                 for key, value in decodingTable.items()
                 if value is not None
                }