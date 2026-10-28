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
        path = str(PurePosixPath(*args)).replace('\\', '/')
        # Check if the path starts with a / because if it does it needs to go.
        if path.startswith('//'):
            path = path[2:]
        elif path.startswith('/'):
            path = path[1:]

        # We also could remove a tailing slash but equality ends up the same so
        # it's not considered important.

        return PurePosixPath.__new__(cls, path)

    def __bool__(self) -> bool:
        return bool(self.parts)

    def __eq__(self, other):
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