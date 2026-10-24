import struct

import extract_msg


if __name__ == '__main__':
    s0 = bytearray(4096)                      # __properties_version1.0
    s0[32:48] = struct.pack('<IIQ', 0x340D0003, 6, 0x40000)
    s0[48:64] = struct.pack('<IIQ', 0x1000001F, 6, 26)
    s0[64:80] = struct.pack("<IIQ", 0x001A001F, 6, 18)
    s0[80:96] = struct.pack('<IIQ', 0x00001102, 6, 32)
    s0[96:112] = struct.pack('<IIQ', 0x0001101F, 6, 16)
    s0[112:128] = struct.pack('<IIQ', 0x00021102, 6, 32)
    s0[128:144] = struct.pack('<IIQ', 0x0003101F, 6, 16)
    s0[144:160] = struct.pack('<IIQ', 0x00041102, 6, 32)
    s0[160:176] = struct.pack('<IIQ', 0x0005101F, 6, 16)
    s0[176:192] = struct.pack('<IIQ', 0x00061102, 6, 32)
    s0[192:208] = struct.pack('<IIQ', 0x0007101F, 6, 16)
    s0[208:224] = struct.pack('<IIQ', 0xF0001102, 6, 32 - 1)
    s0[224:240] = struct.pack('<IIQ', 0xF001101F, 6, 16 - 1)
    s0[240:256] = struct.pack('<IIQ', 0xF003101F, 6, 16)

    s0 = bytes(s0[0:256])
    a = extract_msg.OleWriter()
    a.addEntry('__substg1.0_001A001F', 'IPM.Note'.encode('utf-16-le'))

    binLength = b'\x04\x00\x00\x00\x00\x00\x00\x00' * 4
    bin1 = b'\x00\x01\x02\x03'
    bin2 = b'\x04\x05\x06\x07'
    bin3 = b'\x08\x09\x0A\x0B'
    bin4 = b'\x0C\x0D\x0E\x0F'
    strLength = b'\x0C\x00\x00\x00\x0C\x00\x00\x00\x0A\x00\x00\x00\x08\x00\x00\x00'
    str1 = 'Hello\x00'.encode('utf-16-le')
    str2 = 'World\x00'.encode('utf-16-le')
    str3 = 'It\'s\x00'.encode('utf-16-le')
    str4 = 'Me!\x00'.encode('utf-16-le')
    # Create test streams.
    # Binary tests.
    ## Normal test (success).
    a.addEntry('__substg1.0_00001102', binLength)
    a.addEntry('__substg1.0_00001102-00000000', bin1)
    a.addEntry('__substg1.0_00001102-00000001', bin2)
    a.addEntry('__substg1.0_00001102-00000002', bin3)
    a.addEntry('__substg1.0_00001102-00000003', bin4)

    ## Too few entries test.
    a.addEntry('__substg1.0_00021102', binLength)
    a.addEntry('__substg1.0_00021102-00000000', bin1)
    a.addEntry('__substg1.0_00021102-00000001', bin2)
    a.addEntry('__substg1.0_00021102-00000002', bin3)
    #a.addEntry('__substg1.0_00021102-00000003', bin4)

    ## Missing entry test.
    a.addEntry('__substg1.0_00041102', binLength)
    a.addEntry('__substg1.0_00041102-00000000', bin1)
    a.addEntry('__substg1.0_00041102-00000001', bin2)
    #a.addEntry('__substg1.0_00041102-00000002', bin3)
    a.addEntry('__substg1.0_00041102-00000003', bin4)

    ## Wrong stream length test.
    a.addEntry('__substg1.0_00061102', b'\x03' + binLength[1:])
    a.addEntry('__substg1.0_00061102-00000000', bin1)
    a.addEntry('__substg1.0_00061102-00000001', bin2)
    a.addEntry('__substg1.0_00061102-00000002', bin3)
    a.addEntry('__substg1.0_00061102-00000003', bin4)

    ## Wrong length stream size test.
    a.addEntry('__substg1.0_F0001102', binLength[:-1])
    a.addEntry('__substg1.0_F0001102-00000000', bin1)
    a.addEntry('__substg1.0_F0001102-00000001', bin2)
    a.addEntry('__substg1.0_F0001102-00000002', bin3)
    a.addEntry('__substg1.0_F0001102-00000003', bin4)

    # String tests.
    ## Normal test (success).
    a.addEntry('__substg1.0_0001101F', strLength)
    a.addEntry('__substg1.0_0001101F-00000000', str1)
    a.addEntry('__substg1.0_0001101F-00000001', str2)
    a.addEntry('__substg1.0_0001101F-00000002', str3)
    a.addEntry('__substg1.0_0001101F-00000003', str4)

    ## Too few entries test.
    a.addEntry('__substg1.0_0003101F', strLength)
    a.addEntry('__substg1.0_0003101F-00000000', str1)
    a.addEntry('__substg1.0_0003101F-00000001', str2)
    a.addEntry('__substg1.0_0003101F-00000002', str3)
    #a.addEntry('__substg1.0_0003101F-00000003', str4)

    ## Missing entry test.
    a.addEntry('__substg1.0_0005101F', strLength)
    a.addEntry('__substg1.0_0005101F-00000000', str1)
    a.addEntry('__substg1.0_0005101F-00000001', str2)
    #a.addEntry('__substg1.0_0005101F-00000002', str3)
    a.addEntry('__substg1.0_0005101F-00000003', str4)

    ## Wrong stream size test.
    a.addEntry('__substg1.0_0007101F', b'\x03' + strLength[1:])
    a.addEntry('__substg1.0_0007101F-00000000', str1)
    a.addEntry('__substg1.0_0007101F-00000001', str2)
    a.addEntry('__substg1.0_0007101F-00000002', str3)
    a.addEntry('__substg1.0_0007101F-00000003', str4)

    ## Wrong length stream size test.
    a.addEntry('__substg1.0_F001101F', strLength[:-1])
    a.addEntry('__substg1.0_F001101F-00000000', str1)
    a.addEntry('__substg1.0_F001101F-00000001', str2)
    a.addEntry('__substg1.0_F001101F-00000002', str3)
    a.addEntry('__substg1.0_F001101F-00000003', str4)

    ## No null character test.
    a.addEntry('__substg1.0_F003101F', strLength)
    a.addEntry('__substg1.0_F003101F-00000000', str1[:-2])
    a.addEntry('__substg1.0_F003101F-00000001', str2)
    a.addEntry('__substg1.0_F003101F-00000002', str3)
    a.addEntry('__substg1.0_F003101F-00000003', str4)

    # Set body for stdout test.
    a.addEntry('__substg1.0_1000001F', 'Hello World!'.encode('utf-16-le'))
    a.addEntry('__properties_version1.0', s0)
    a.write('multi-stream.msg')
