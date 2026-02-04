//go:build ignore
// +build ignore

package main

import (
	"fmt"
)

// decodeSingleByteString converts bytes to string treating each byte as a rune (ISO-8859-1)
func decodeSingleByteString(data []byte) string {
	runes := make([]rune, len(data))
	for i, b := range data {
		runes[i] = rune(b)
	}
	return string(runes)
}

// encodeSingleByteString converts string to bytes treating each char as a byte (ISO-8859-1)
func encodeSingleByteString(s string) []byte {
	data := make([]byte, 0, len(s))
	for _, r := range s {
		if r <= 0xFF {
			data = append(data, byte(r))
		} else {
			// Replace non-ISO-8859-1 chars with '?'
			data = append(data, '?')
		}
	}
	return data
}

func main() {
	// Test with multipart boundary containing various bytes
	original := []byte("------WebKitFormBoundary\r\nContent-Disposition: form-data; name=\"testfile\"; filename=\"hello.txt\"")

	fmt.Printf("Original bytes: len=%d\n", len(original))
	fmt.Printf("First 30 bytes: %q\n", string(original[:30]))

	// Decode to string
	decoded := decodeSingleByteString(original)
	fmt.Printf("Decoded string: len(bytes)=%d, len(runes)=%d\n", len(decoded), len([]rune(decoded)))

	// Encode back to bytes
	encoded := encodeSingleByteString(decoded)
	fmt.Printf("Encoded bytes: len=%d\n", len(encoded))
	fmt.Printf("First 30 bytes: %q\n", string(encoded[:30]))

	// Check if they match
	match := len(original) == len(encoded)
	if match {
		for i := range original {
			if original[i] != encoded[i] {
				match = false
				fmt.Printf("Mismatch at position %d: original=%d, encoded=%d\n", i, original[i], encoded[i])
				break
			}
		}
	}
	fmt.Printf("Match: %v\n", match)

	// Now test with Position-like access
	fmt.Println("\n--- Position test ---")
	// Simulating: Position = 42 (0-based, expecting "Content-Disposition")
	pos := 42
	data := encoded[pos : pos+19]
	fmt.Printf("Position %d, reading 19 bytes: %q\n", pos, string(data))
}
