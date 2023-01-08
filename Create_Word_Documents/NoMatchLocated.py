class NoMatchFound(Exception):
    def __init__(self, message="Match was not found."):
        self.message = message
        super().__init__(self.message)