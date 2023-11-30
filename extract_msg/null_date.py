__all__ = [
    'NullDate'
]


import datetime

from typing import Optional


class NullDate(datetime.datetime):
    """
    Version of datetime.datetime intended to represent a value of NULL.

    Some properties use different values for a null date, and those need to be
    differentiated when packing the data back into an MSG file, which is why
    this class exists. Comparisons between NullDate instances will always say
    the two dates are equal.

    :attribute filetime: An optional value that can be set to the filetime a
        null date should convert back to.
    """

    filetime: Optional[int] = None

    def __eq__(self, other) -> bool:
        if isinstance(other, NullDate):
            return True
        return super.__eq__(self, other)

    def __ne__(self, other) -> bool:
        if isinstance(other, NullDate):
            return False
        return super.__eq__(self, other)