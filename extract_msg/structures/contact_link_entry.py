__all__ = [
    'ContactLinkEntry',
]


from typing import List

from ._helpers import BytesReader
from ..constants import st
from .entry_id import EntryID


class ContactLinkEntry:
    entries: List[EntryID]

    def __init__(self, data: bytes):
        # My experience with this data almost entirely doesn't match the
        # documentation, so I'm just going to do what I see and not what I'm
        # told.
        reader = BytesReader(data)
        count = reader.readUnsignedInt()
        # Ignore this field.
        reader.read(4)
        self.entries = []
        for _ in range(count):
            size = reader.readUnsignedInt()
            self.entries.append(EntryID.autoCreate(reader.read(size)))
            if (size & 3) != 0:
                reader.read(4 - (size & 3))

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        ret = st.ST_LE_UI32.pack(len(self.entries))

        # Need to handle the data before hand.
        data = b''
        for entry in self.entries:
            entryData = entry.toBytes()
            # Size goes before data.
            data += st.ST_LE_UI32.pack(edLen := len(entryData))
            data += entryData
            # Handle padding.
            if edLen & 3:
                data += b'\x00' * (4 - edLen)

        ret += st.ST_LE_UI32.pack(len(data))
        ret += data

        return ret
