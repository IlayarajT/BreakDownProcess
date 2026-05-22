import re


class HexEscapeConverter:
    """
    Converts non-ASCII characters to 6-digit hexadecimal escape sequences
    and converts those escape sequences back to Unicode characters.

    Behavior is identical to the original implementation.
    """

    def __init__(self):
        # Matches 6-digit hexadecimal escape sequences like \x0001f60a
        self.hex_escape_pattern = re.compile(r'\\x([0-9a-fA-F]{6})')

    def char_to_hex_escape(self, char):
        """
        Convert a single character to its 6-digit hexadecimal escape sequence.
        """
        code_point = ord(char)
        return f'\\x{code_point:06x}'

    def convert_non_ascii_to_hex_escapes(self, text):
        """
        Convert non-ASCII characters in a string to
        6-digit hexadecimal escape sequences.
        """
        result = []
        for char in text:
            if ord(char) > 127:
                result.append(self.char_to_hex_escape(char))
            else:
                result.append(char)
        return ''.join(result)

    def hex_escape_to_char(self, match):
        """
        Convert a 6-digit hexadecimal escape sequence
        to its corresponding Unicode character.
        """
        code_point = int(match.group(1), 16)
        return chr(code_point)

    def convert_hex_escapes_to_utf8(self, text):
        """
        Convert 6-digit hexadecimal escape sequences
        in a string back to Unicode characters.
        """
        return self.hex_escape_pattern.sub(self.hex_escape_to_char, text)
