package asp

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

// Include regex: <!--#include file="path"--> or <!--#include virtual="/path"-->
// Case insensitive, handles whitespace
var includeRegex = regexp.MustCompile(`(?i)<!--\s*#include\s+(file|virtual)\s*=\s*"([^"]+)"\s*-->`)

// ReadFile reads a file from the file system
func ReadFile(path string) ([]byte, error) {
	return os.ReadFile(path)
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
		includedContent, err := os.ReadFile(targetPath)
		if err != nil {
			// In ASP, missing include usually throws error
			// We can append an error message comment or fail
			// Faiing is safer/standard
			return "", fmt.Errorf("failed to include file '%s': %w", incPath, err)
		}

		// Recursively resolve includes in the included content
		resolvedIncluded, err := ResolveIncludes(string(includedContent), targetPath, rootDir, visited)
		if err != nil {
			return "", err
		}

		sb.WriteString(resolvedIncluded)

		lastIndex = match[1]
	}

	// Append remaining content
	sb.WriteString(content[lastIndex:])

	return sb.String(), nil
}