/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package asp

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"unicode/utf16"
)

// Include regex: <!--#include file="path"--> or <!--#include virtual="/path"-->
// Case insensitive, handles whitespace
var includeRegex = regexp.MustCompile(`(?i)<!--\s*#include\s+(file|virtual)\s*=\s*"([^"]+)"\s*-->`)

// METADATA directive regex: <!--METADATA TYPE="TypeLib" ... -->
// IIS uses these to import type library constants; we strip them since constants are built-in.
var metadataRegex = regexp.MustCompile(`(?is)<!--\s*METADATA\s+[^>]*?-->`)

// ReadFile reads a file from the file system
func ReadFile(path string) ([]byte, error) {
	return os.ReadFile(path)
}

// ReadFileText reads a text file and converts it to UTF-8, honoring BOM and common UTF-16 encodings.
func ReadFileText(path string) (string, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return "", err
	}
	return decodeTextToUTF8(data), nil
}

// ResolveIncludes processes SSI include directives recursively
// content: The content of the file
// currentFile: The absolute path of the file being processed (to resolve 'file' includes)
// rootDir: The web root directory (to resolve 'virtual' includes)
// visited: Map to detect circular dependencies
func ResolveIncludes(content, currentFile, rootDir string, visited map[string]bool) (string, error) {
	if visited == nil {
		visited = make(map[string]bool)
	}

	// Normalize paths
	currentFile = filepath.Clean(currentFile)
	rootDir = filepath.Clean(rootDir)

	// Check circular dependency
	if visited[currentFile] {
		return "", fmt.Errorf("circular include detected: %s", currentFile)
	}
	visited[currentFile] = true
	defer delete(visited, currentFile) // Allow re-visiting in other branches, but usually includes are one-pass

	// Strip <!--METADATA --> directives (IIS type library references)
	content = metadataRegex.ReplaceAllString(content, "")

	// Find all matches
	matches := includeRegex.FindAllStringSubmatchIndex(content, -1)

	if len(matches) == 0 {
		return content, nil
	}

	var sb strings.Builder
	lastIndex := 0

	for _, match := range matches {
		// Append content before the match
		sb.WriteString(content[lastIndex:match[0]])

		// Extract type (file/virtual) and path
		incType := strings.ToLower(content[match[2]:match[3]])
		incPath := content[match[4]:match[5]]

		// Resolve path
		var targetPath string
		if incType == "virtual" {
			// Virtual: relative to rootDir
			// Remove leading slash
			cleanIncPath := strings.TrimPrefix(strings.ReplaceAll(incPath, "\\", "/"), "/")
			targetPath = filepath.Join(rootDir, cleanIncPath)
		} else {
			// File: relative to current directory
			currentDir := filepath.Dir(currentFile)
			targetPath = filepath.Join(currentDir, incPath)
		}

		// Read included file
		includedContent, err := ReadFileText(targetPath)
		if err != nil {
			// In ASP, missing include usually throws error
			// We can append an error message comment or fail
			// Faiing is safer/standard
			return "", fmt.Errorf("failed to include file '%s': %w", incPath, err)
		}

		// Recursively resolve includes in the included content
		resolvedIncluded, err := ResolveIncludes(includedContent, targetPath, rootDir, visited)
		if err != nil {
			return "", err
		}

		// Preserve statement boundaries when includes sit between ASP blocks (e.g. %>...<%).
		before := content[:match[0]]
		after := content[match[1]:]
		needsLeadBreak := strings.HasSuffix(before, "%>")
		needsTrailBreak := strings.HasPrefix(after, "<%")

		if needsLeadBreak && resolvedIncluded != "" && !strings.HasPrefix(resolvedIncluded, "\n") && !strings.HasPrefix(resolvedIncluded, "\r") {
			sb.WriteString("\n")
		}
		sb.WriteString(resolvedIncluded)
		if needsTrailBreak && resolvedIncluded != "" && !strings.HasSuffix(resolvedIncluded, "\n") && !strings.HasSuffix(resolvedIncluded, "\r") {
			sb.WriteString("\n")
		}

		lastIndex = match[1]
	}

	// Append remaining content
	sb.WriteString(content[lastIndex:])

	return sb.String(), nil
}

// decodeTextToUTF8 converts raw text bytes to UTF-8, removing BOMs and handling UTF-16 LE/BE.
func decodeTextToUTF8(data []byte) string {
	if len(data) >= 3 && bytes.Equal(data[:3], []byte{0xEF, 0xBB, 0xBF}) {
		return string(data[3:])
	}

	if len(data) >= 2 {
		switch {
		case data[0] == 0xFF && data[1] == 0xFE:
			return decodeUTF16(data[2:], binary.LittleEndian)
		case data[0] == 0xFE && data[1] == 0xFF:
			return decodeUTF16(data[2:], binary.BigEndian)
		}
	}

	zeroEven, zeroOdd := 0, 0
	for i := 0; i < len(data); i++ {
		if data[i] == 0 {
			if i%2 == 0 {
				zeroEven++
			} else {
				zeroOdd++
			}
		}
	}

	if zeroOdd > len(data)/4 && zeroOdd >= zeroEven {
		return decodeUTF16(data, binary.LittleEndian)
	}
	if zeroEven > len(data)/4 && zeroEven > zeroOdd {
		return decodeUTF16(data, binary.BigEndian)
	}

	return string(data)
}

func decodeUTF16(data []byte, order binary.ByteOrder) string {
	if len(data) == 0 {
		return ""
	}
	if len(data)%2 != 0 {
		data = data[:len(data)-1]
	}

	u16 := make([]uint16, len(data)/2)
	for i := range u16 {
		u16[i] = order.Uint16(data[i*2:])
	}
	return string(utf16.Decode(u16))
}
