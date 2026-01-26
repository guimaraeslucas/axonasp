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
package server

import (
	"fmt"
	"regexp"
	"strings"
)

// G3REGEXP implements the RegExp object for pattern matching and manipulation
// Follows VBScript RegExp behavior with Go's regexp engine underneath
type G3REGEXP struct {
	pattern     string         // The pattern string
	ignoreCase  bool           // Case-insensitive matching
	global      bool           // Match all occurrences
	multiline   bool           // Multi-line mode (^ and $ match line boundaries)
	compiled    *regexp.Regexp // Compiled regex
	lastMatches []*RegExpMatch // Last execution matches
	err         string         // Last error message
}

// RegExpMatch represents a single match result
type RegExpMatch struct {
	Value  string // Matched text
	Index  int64  // 0-based position in string
	Length int64  // Length of match
}

// RegExpMatches represents a collection of matches
type RegExpMatches struct {
	matches []*RegExpMatch
	count   int64
}

// GetProperty gets a property from the RegExp object
func (r *G3REGEXP) GetProperty(name string) any {
	switch strings.ToLower(name) {
	case "pattern":
		return r.pattern
	case "ignorecase":
		return r.ignoreCase
	case "global":
		return r.global
	case "multiline":
		return r.multiline
	case "source":
		return r.pattern
	default:
		return nil
	}
}

// SetProperty sets a property on the RegExp object
func (r *G3REGEXP) SetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "pattern":
		if s, ok := value.(string); ok {
			r.pattern = s
			r.compilePattern()
		}
	case "ignorecase":
		r.ignoreCase = toTruthy(value)
		r.compilePattern()
	case "global":
		r.global = toTruthy(value)
		r.compilePattern()
	case "multiline":
		r.multiline = toTruthy(value)
		r.compilePattern()
	case "source":
		if s, ok := value.(string); ok {
			r.pattern = s
			r.compilePattern()
		}
	}
	return nil
}

// CallMethod calls a method on the RegExp object
func (r *G3REGEXP) CallMethod(name string, args ...interface{}) (interface{}, error) {
	nameLower := strings.ToLower(name)

	switch nameLower {
	// Methods
	case "execute":
		if len(args) > 0 {
			return r.Execute(fmt.Sprintf("%v", args[0])), nil
		}
		return nil, nil
	case "test":
		if len(args) > 0 {
			return r.Test(fmt.Sprintf("%v", args[0])), nil
		}
		return false, nil
	case "replace":
		if len(args) >= 2 {
			return r.Replace(fmt.Sprintf("%v", args[0]), fmt.Sprintf("%v", args[1])), nil
		}
		return "", nil

	// Property getters (direct call)
	case "getpattern":
		return r.pattern, nil
	case "getignorecase":
		return r.ignoreCase, nil
	case "getglobal":
		return r.global, nil
	case "getmultiline":
		return r.multiline, nil
	case "getsource":
		return r.pattern, nil

	// Property setters (direct call)
	case "setpattern":
		if len(args) > 0 {
			r.pattern = fmt.Sprintf("%v", args[0])
			r.compilePattern()
		}
		return nil, nil
	case "setignorecase":
		if len(args) > 0 {
			r.ignoreCase = toTruthy(args[0])
			r.compilePattern()
		}
		return nil, nil
	case "setglobal":
		if len(args) > 0 {
			r.global = toTruthy(args[0])
			r.compilePattern()
		}
		return nil, nil
	case "setmultiline":
		if len(args) > 0 {
			r.multiline = toTruthy(args[0])
			r.compilePattern()
		}
		return nil, nil
	case "setsource":
		if len(args) > 0 {
			r.pattern = fmt.Sprintf("%v", args[0])
			r.compilePattern()
		}
		return nil, nil

	// Generic SetProperty/GetProperty handlers (when called as methods)
	case "setproperty", "set":
		// setproperty(propertyName, value)
		if len(args) >= 2 {
			propName := strings.ToLower(fmt.Sprintf("%v", args[0]))
			propValue := args[1]

			switch propName {
			case "pattern", "source":
				r.pattern = fmt.Sprintf("%v", propValue)
				r.compilePattern()
			case "ignorecase":
				r.ignoreCase = toTruthy(propValue)
				r.compilePattern()
			case "global":
				r.global = toTruthy(propValue)
				r.compilePattern()
			case "multiline":
				r.multiline = toTruthy(propValue)
				r.compilePattern()
			}
		}
		return nil, nil

	case "getproperty", "get":
		// getproperty(propertyName)
		if len(args) > 0 {
			propName := strings.ToLower(fmt.Sprintf("%v", args[0]))
			switch propName {
			case "pattern", "source":
				return r.pattern, nil
			case "ignorecase":
				return r.ignoreCase, nil
			case "global":
				return r.global, nil
			case "multiline":
				return r.multiline, nil
			}
		}
		return nil, nil

	// Property access without Get/Set prefix (for direct method calls)
	case "pattern":
		if len(args) > 0 {
			r.pattern = fmt.Sprintf("%v", args[0])
			r.compilePattern()
			return nil, nil
		}
		return r.pattern, nil
	case "ignorecase":
		if len(args) > 0 {
			r.ignoreCase = toTruthy(args[0])
			r.compilePattern()
			return nil, nil
		}
		return r.ignoreCase, nil
	case "global":
		if len(args) > 0 {
			r.global = toTruthy(args[0])
			r.compilePattern()
			return nil, nil
		}
		return r.global, nil
	case "multiline":
		if len(args) > 0 {
			r.multiline = toTruthy(args[0])
			r.compilePattern()
			return nil, nil
		}
		return r.multiline, nil
	case "source":
		if len(args) > 0 {
			r.pattern = fmt.Sprintf("%v", args[0])
			r.compilePattern()
			return nil, nil
		}
		return r.pattern, nil

	default:
		return nil, nil
	}
}

// compilePattern compiles the pattern with appropriate flags
func (r *G3REGEXP) compilePattern() {
	if r.pattern == "" {
		r.compiled = nil
		r.err = ""
		return
	}

	pattern := r.pattern

	// Handle multiline mode
	if r.multiline {
		pattern = "(?m)" + pattern
	}

	// Go's regexp is always case-sensitive by default
	// For case-insensitive, we need to add (?i) flag
	if r.ignoreCase {
		pattern = "(?i)" + pattern
	}

	compiled, err := regexp.Compile(pattern)
	if err != nil {
		r.compiled = nil
		r.err = err.Error()
		return
	}

	r.compiled = compiled
	r.err = ""
}

// Execute searches for all matches of the pattern in the string
// Returns a RegExpMatches collection (or similar structure)
func (r *G3REGEXP) Execute(inputStr string) interface{} {
	if r.compiled == nil {
		if r.pattern == "" {
			r.err = "Pattern not set"
		}
		return &RegExpMatchesCollection{
			matches: []*RegExpMatch{},
			count:   0,
		}
	}

	var matches []*RegExpMatch

	if r.global {
		// Find all matches
		allMatches := r.compiled.FindAllStringIndex(inputStr, -1)
		for _, match := range allMatches {
			if len(match) >= 2 {
				matchValue := inputStr[match[0]:match[1]]
				matches = append(matches, &RegExpMatch{
					Value:  matchValue,
					Index:  int64(match[0]),
					Length: int64(len(matchValue)),
				})
			}
		}
	} else {
		// Find first match only
		matchIdx := r.compiled.FindStringIndex(inputStr)
		if matchIdx != nil && len(matchIdx) >= 2 {
			matchValue := inputStr[matchIdx[0]:matchIdx[1]]
			matches = append(matches, &RegExpMatch{
				Value:  matchValue,
				Index:  int64(matchIdx[0]),
				Length: int64(len(matchValue)),
			})
		}
	}

	r.lastMatches = matches

	return &RegExpMatchesCollection{
		matches: matches,
		count:   int64(len(matches)),
	}
}

// Test checks if the pattern matches the string
// Returns true if at least one match is found
func (r *G3REGEXP) Test(inputStr string) bool {
	if r.compiled == nil {
		return false
	}

	return r.compiled.MatchString(inputStr)
}

// Replace replaces matched patterns with replacement text
// Supports $& (entire match), $` (before match), $' (after match)
func (r *G3REGEXP) Replace(inputStr string, replacement string) string {
	if r.compiled == nil {
		return inputStr
	}

	// Convert VBScript replacement syntax to Go syntax
	// $& -> $0 (entire match)
	// $` -> before match text
	// $' -> after match text
	// For simplicity, we'll process $& and handle others manually

	result := ""

	if r.global {
		// Replace all matches
		lastIndex := 0
		matches := r.compiled.FindAllStringSubmatchIndex(inputStr, -1)

		for _, match := range matches {
			// Append text before match
			result += inputStr[lastIndex:match[0]]

			// Process replacement string
			replacedText := r.processReplacement(replacement, inputStr, match)
			result += replacedText

			lastIndex = match[1]
		}

		// Append remaining text
		result += inputStr[lastIndex:]
		return result
	} else {
		// Replace first match only
		matchIdx := r.compiled.FindStringSubmatchIndex(inputStr)
		if matchIdx != nil {
			// Append text before match
			result := inputStr[:matchIdx[0]]

			// Process replacement string
			replacedText := r.processReplacement(replacement, inputStr, matchIdx)
			result += replacedText

			// Append text after match
			result += inputStr[matchIdx[1]:]
			return result
		}
	}

	return inputStr
}

// processReplacement processes the replacement string with special sequences
func (r *G3REGEXP) processReplacement(replacement string, inputStr string, matchIdx []int) string {
	if len(matchIdx) < 2 {
		return replacement
	}

	result := replacement
	matchText := inputStr[matchIdx[0]:matchIdx[1]]

	// Replace special sequences
	// $& = entire match
	result = strings.ReplaceAll(result, "$&", matchText)

	// $` = text before match
	if matchIdx[0] > 0 {
		result = strings.ReplaceAll(result, "$`", inputStr[:matchIdx[0]])
	} else {
		result = strings.ReplaceAll(result, "$`", "")
	}

	// $' = text after match
	if matchIdx[1] < len(inputStr) {
		result = strings.ReplaceAll(result, "$'", inputStr[matchIdx[1]:])
	} else {
		result = strings.ReplaceAll(result, "$'", "")
	}

	// Remove any unmatched $ sequences
	result = regexp.MustCompile(`\$\d`).ReplaceAllString(result, "")

	return result
}

// RegExpMatchesCollection represents a collection of matches (returned by Execute)
type RegExpMatchesCollection struct {
	matches []*RegExpMatch
	count   int64
}

func (m *RegExpMatchesCollection) GetName() string {
	return "IMatchCollection2"
}

// GetProperty gets a property from matches collection
func (m *RegExpMatchesCollection) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "count":
		return m.count
	case "length":
		return m.count
	case "item":
		return m // Return self for indexing
	default:
		return nil
	}
}

// SetProperty sets a property (no-op for collection)
func (m *RegExpMatchesCollection) SetProperty(name string, value interface{}) error {
	return nil
}

// CallMethod for collection - primarily Item access (with error return for interface compatibility)
func (m *RegExpMatchesCollection) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch strings.ToLower(name) {
	case "item":
		if len(args) > 0 {
			idx, ok := toInt64(args[0])
			if ok && idx >= 0 && idx < m.count {
				return m.matches[idx], nil
			}
		}
		return nil, fmt.Errorf("invalid index")
	case "count":
		return m.count, nil
	default:
		// Default property access matches(0)
		if len(args) > 0 {
			idx, ok := toInt64(args[0])
			if ok && idx >= 0 && idx < m.count {
				return m.matches[idx], nil
			}
		}
		return nil, fmt.Errorf("method not found: %s", name)
	}
}

// RegExpMatch Methods
func (m *RegExpMatch) GetName() string {
	return "IMatch2"
}

// GetProperty on a single match
func (m *RegExpMatch) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "value":
		return m.Value
	case "firstindex":
		return m.Index // VBScript uses 0-based indexing for FirstIndex
	case "index":
		return m.Index // Alias
	case "length":
		return m.Length
	default:
		return nil
	}
}

// SetProperty on a single match (read-only)
func (m *RegExpMatch) SetProperty(name string, value interface{}) error {
	return nil
}

// CallMethod on a single match (with error return for interface compatibility)
func (m *RegExpMatch) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch strings.ToLower(name) {
	case "value":
		return m.Value, nil
	case "firstindex":
		return m.Index, nil
	case "index":
		return m.Index, nil
	case "length":
		return m.Length, nil
	default:
		// Default property is Value
		return m.Value, nil
	}
}

// Helper function to convert value to int64
func toInt64(value interface{}) (int64, bool) {
	switch v := value.(type) {
	case int:
		return int64(v), true
	case int32:
		return int64(v), true
	case int64:
		return v, true
	case float64:
		return int64(v), true
	case string:
		if i, err := fmt.Sscanf(v, "%d"); err == nil {
			return int64(i), true
		}
	}
	return 0, false
}

// Helper function to check truthiness (local version)
func toTruthy(value interface{}) bool {
	if value == nil {
		return false
	}
	if b, ok := value.(bool); ok {
		return b
	}
	if i, ok := value.(int); ok {
		return i != 0
	}
	if f, ok := value.(float64); ok {
		return f != 0
	}
	if s, ok := value.(string); ok {
		lower := strings.ToLower(s)
		return lower != "" && lower != "false" && lower != "0"
	}
	return false
}
