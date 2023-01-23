class RTFTokenizer:
    def __init__(self, data : bytes = None):

        # Feed the data to our parser if provided. Feeding will clear all of the
        # existing data.
        if data:
            self.feed(data)

    def feed(self, data : bytes) -> None:
        pass
