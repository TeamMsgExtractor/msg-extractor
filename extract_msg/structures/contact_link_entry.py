__all__ = [
    'ContactLinkEntry',
]


from typing import List

from ._helpers import BytesReader
from .entry_id import AddressBookEntryID


class ContactLinkEntry:
    entries : List[AddressBookEntryID]

    def __init__(self, data : bytes):
        reader = BytesReader(data)
        count = reader.readUnsignedInt()
        reader.read(4)
        remaining = reader.read()
        self.entries = []
        for _ in range(count):
            idStruct = AddressBookEntryID(remaining)
            remaining = remaining[idStruct.position:]