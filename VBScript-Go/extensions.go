package vbscript

import (
	"strings"
)

// GetChar returns the character at a given position in a string, or rune(0) if out of bounds
func GetChar(str string, pos int) rune {
	if pos < 0 || pos >= len(str) {
		return rune(0)
	}
	return []rune(str)[pos]
}

// GetCharCode returns the numeric code of a character at a given position
func GetCharCode(str string, pos int) int {
	if pos < 0 || pos >= len(str) {
		return 0
	}
	return int([]rune(str)[pos])
}

// Slice returns a substring from start to end position
// Supports negative indices like JavaScript slice
// end is exclusive
func Slice(str string, start, end int) string {
	runes := []rune(str)
	length := len(runes)

	// Convert negative indices
	from := start
	if from < 0 {
		from = length + from
		if from < 0 {
			from = 0
		}
	}
	if from > length {
		from = length
	}

	to := end
	if to < 0 {
		to = length + to
		if to < 0 {
			to = 0
		}
	}
	if to > length {
		to = length
	}

	// Ensure from <= to
	if from > to {
		return ""
	}

	return string(runes[from:to])
}

// CIEquals performs case-insensitive string comparison
func CIEquals(s1, s2 string) bool {
	return strings.EqualFold(s1, s2)
}

// StringToLower converts a string to lowercase (using English locale)
func StringToLower(s string) string {
	return strings.ToLower(s)
}

// StringToUpper converts a string to uppercase (using English locale)
func StringToUpper(s string) string {
	return strings.ToUpper(s)
}
