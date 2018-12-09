import logging
import struct

from extract_msg import constants
from extract_msg.utils import divide  # , round_up

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


# TODO move this function to utils.py:
def round_up(inp, mult):
    """
    Rounds :param inp: up to the nearest multiple of :param mult:.
    """
    return inp + (mult - inp) % mult


# Temporary class code to make references like `constants.CONSTANT` work:
class constants(object):
    # Structs used by named.py
    STNP_NAM = struct.Struct('<i')
    STNP_ENT = struct.Struct('<IHH')
    # TODO move these out of the class and into constants.py:


# TODO
class Named(object):
    def __init__(self, msg):  # Temporarily uses the Message instance as the input
        super(Named, self).__init__()
        guid_stream = msg._getStream('__nameid_version1.0/__substg1.0_00020102')
        entry_stream = msg._getStream('__nameid_version1.0/__substg1.0_00030102')
        names_stream = msg._getStream('__nameid_version1.0/__substg1.0_00040102')
        gl = len(guid_stream)
        el = len(entry_stream)
        nl = len(names_stream)
        # TODO guid stream parsing
        # TODO entry_stream parsing
        entries = []
        for x in divide(entry_stream, 8):
            tmp = constants.STNP_ENT.unpack(x)
            entries.append({
                'id': tmp[0],
                'pid': tmp[2],
                'guid': tmp[1] >> 1,
                'pkind': tmp[1] & 1,
            })
        names = []
        pos = 0
        while pos < nl:
            l = constants.STNP_NAM.unpack(names_stream[pos:pos + 4])[0]
            pos += 4
            names.append(names_stream[pos:pos + l].decode('utf_16_le'))
            pos += round_up(l, 4)
