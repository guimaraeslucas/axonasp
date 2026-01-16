package vbscript

// IsLineTerminator checks if a character is a line terminator
func IsLineTerminator(c rune) bool {
	return c == '\n' || c == '\r' || c == ':'
}

// IsNewLine checks if a character is a new line character
func IsNewLine(c rune) bool {
	return c == '\n' || c == '\r'
}

// IsDecDigit checks if a character is a decimal digit (0-9)
func IsDecDigit(c rune) bool {
	return c >= '0' && c <= '9'
}

// IsHexDigit checks if a character is a hexadecimal digit (0-9, A-F, a-f)
func IsHexDigit(c rune) bool {
	return IsDecDigit(c) ||
		(c >= 'a' && c <= 'f') ||
		(c >= 'A' && c <= 'F')
}

// IsOctDigit checks if a character is an octal digit (0-7)
func IsOctDigit(c rune) bool {
	return c >= '0' && c <= '7'
}

// IsIdentifierStart checks if a character can start an identifier
func IsIdentifierStart(c rune) bool {
	return (c >= 'A' && c <= 'Z') ||
		(c >= 'a' && c <= 'z')
}

// IsIdentifier checks if a character is valid in an identifier
func IsIdentifier(c rune) bool {
	return IsIdentifierStart(c) ||
		IsDecDigit(c) ||
		c == '_'
}

// IsWhiteSpace checks if a character is whitespace
func IsWhiteSpace(c rune) bool {
	return c == 0x20 || c == 0x09 ||
		c == 0x0B || c == 0x0C
}

// IsExtendedIdentifier checks if a character is valid in an extended identifier
func IsExtendedIdentifier(c rune) bool {
	return !IsNewLine(c) && c != ']' && c >= 0 && c <= 0xff
}

// CharEquals compares two characters case-insensitively
func CharEquals(a, b rune) bool {
	return runeToUpper(a) == runeToUpper(b)
}

// runeToUpper converts a rune to uppercase
func runeToUpper(c rune) rune {
	if c >= 'a' && c <= 'z' {
		return c - 32
	}
	return c
}
