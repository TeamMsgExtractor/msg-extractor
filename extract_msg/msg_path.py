from __future__ import annotations

__all__ = [
    'MSGPath',
]

from pathlib import PurePosixPath
from typing import Any


class MSGPath(PurePosixPath):
    def __new__(cls, *args: Any):
        # Ensure we handle any old-style msg path types as arguments.
        newArgs = []
        for item in args:
            if isinstance(item, tuple) or isinstance(item, list):
                newArgs += item
            else:
                newArgs.append(item)

        # We need to handle ensuring that the path is normalized.
        # First, normalize the slashes.
        path = str(PurePosixPath(*newArgs)).replace('\\', '/')
        # Check if the path starts with a / because if it does it needs to go.
        if path.startswith('//'):
            path = path[2:]
        elif path.startswith('/'):
            path = path[1:]

        # We also could remove a tailing slash but equality ends up the same so
        # it's not considered important.

        # Now, we have the path as a string and are about to be done. However
        # our second to last step is to check for invalid characters. If we
        # find any, throw a hard error.
        if '!' in path or ':' in path:
            raise ValueError('Illegal character ("!" or ":") found in MSG path.')

        # Our last one will require us to analyze the parts of the path. As
        # such, convert the path first then check.
        pp = PurePosixPath.__new__(cls, path)

        # Make sure each segment is not more than 31 characters long.
        if any(len(x) > 31 for x in pp.parts):
            raise ValueError('Path segments must not be greater than 31 characters.')

        # All good.
        return pp

    def __bool__(self) -> bool:
        return bool(self.parts)

    def __eq__(self, other) -> bool:
        return str(self).lower() == str(other).lower()

    def __getitem__(self, index: int) -> str:
            return self.parts[index]

    def __len__(self) -> int:
        return len(self.parts)

    def __str__(self) -> str:
        # Implemented for safety to ensure consistency across versions.
        ret = super().__str__()
        if ret.endswith('/'):
            return ret[:-1]
        return ret

    # I would have used __add__ for this but that might obfuscate wrong usage.
    def add(self, tail: str) -> MSGPath:
        """
        Adds the tail to the last item in the path, without a separator.
        """
        return self.__class__(str(self) + tail)