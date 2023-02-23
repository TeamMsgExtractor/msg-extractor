from typing import Optional

from . import RTFTokenizer


class RTFParser(RTFTokenizer):
    """
    Extension of the RTFTokenizer class which handles advanced parsing including
    grouping, determining if the data is in the correct order, etc.
    """

    def __init__(self, data : Optional[bytes] = None):
        super().__init__(data)

    def feed(self, data : bytes) -> None:
        # Backup our tokens, since we need to be more careful with them.
        oldTokens = self.tokens

        # First feed the data to the superclass.
        super().feed(data)

        # TODO
