class TextSegment:
    def __init__(self, content):
        self.content = content
        self.lexemes = []

    def add_token(self, token):
        self.lexemes.append(token)