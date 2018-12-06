from extract_msg import constants
from extract_msg.debug import debug, logger

# TODO move this function to utils:
def round_up(inp, mult):
    """
    Rounds :param inp: up to the nearest multiple of :param mult:.
    """
    return inp + (mult - inp) % mult



# TODO
class Named(object):
    def __init__(self, msg): # Temporarily uses the Message instance as the input
        super(Named, self).__init__()
        guid_stream = msg._getStream('__nameid_version1.0/__substg1.0_00020102')
        entry_stream = msg._getStream('__nameid_version1.0/__substg1.0_00030102')
        names_stream = msg._getStream('__nameid_version1.0/__substg1.0_00040102')
        gl = len(guid_stream)
        el = len(entry_stream)
        nl = len(names_stream)
        # TODO guid stream parsing
        # TODO entry_stream parsing
        names = []
        pos = 0
        while pos < nl:
            l = constants.STNP_NAM.unpack(names_stream[pos:pos+4])[0]
            pos += 4
            names.append(names_stream[pos:pos+l].decode('utf_16_le'))
            pos += round_up(l, 4)
