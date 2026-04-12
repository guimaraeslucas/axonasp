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
package axonvm

import (
	"fmt"
	"regexp"
	"strings"

	"g3pix.com.br/axonasp/vbscript"
)

// regExpNativeObject stores one VBScript RegExp runtime instance.
type regExpNativeObject struct {
	pattern    string
	ignoreCase bool
	global     bool
	multiLine  bool
	compiled   *regexp.Regexp
	lastError  string
}

// regExpMatch represents a single match result.
type regExpMatch struct {
	value      string
	index      int
	length     int
	subMatches *regExpSubMatches
}

// regExpMatchesCollection represents a collection of matches.
type regExpMatchesCollection struct {
	matches []*regExpMatch
}

// regExpSubMatches represents captured groups in a match.
type regExpSubMatches struct {
	values []string
}

// regExpSubMatchValue stores one captured group value object.
type regExpSubMatchValue struct {
	value string
}

// newRegExpObject allocates one RegExp native object.
func (vm *VM) newRegExpObject() Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.regExpItems[objID] = &regExpNativeObject{
		global:     false,
		ignoreCase: false,
		multiLine:  false,
	}
	return Value{Type: VTNativeObject, Num: objID}
}

// dispatchRegExpMethod routes RegExp method calls.
func (vm *VM) dispatchRegExpMethod(objID int64, member string, args []Value) (Value, bool) {
	re, exists := vm.regExpItems[objID]
	if exists {
		return vm.dispatchRegExpObjectMethod(re, member, args), true
	}

	if matches, exists := vm.regExpMatchesCollectionItems[objID]; exists {
		return vm.dispatchRegExpMatchesCollectionMethod(matches, member, args), true
	}

	if match, exists := vm.regExpMatchItems[objID]; exists {
		return vm.dispatchRegExpMatchMethod(match, member, args), true
	}

	if subMatches, exists := vm.regExpSubMatchesItems[objID]; exists {
		return vm.dispatchRegExpSubMatchesMethod(subMatches, member, args), true
	}

	if subMatchValue, exists := vm.regExpSubMatchValueItems[objID]; exists {
		return vm.dispatchRegExpSubMatchValueMethod(subMatchValue, member, args), true
	}

	return Value{Type: VTEmpty}, false
}

// dispatchRegExpPropertyGet resolves RegExp property reads.
func (vm *VM) dispatchRegExpPropertyGet(objID int64, member string) (Value, bool) {
	if re, exists := vm.regExpItems[objID]; exists {
		return vm.dispatchRegExpObjectPropertyGet(re, member), true
	}

	if matches, exists := vm.regExpMatchesCollectionItems[objID]; exists {
		return vm.dispatchRegExpMatchesCollectionPropertyGet(matches, member), true
	}

	if match, exists := vm.regExpMatchItems[objID]; exists {
		return vm.dispatchRegExpMatchPropertyGet(match, member), true
	}

	if subMatches, exists := vm.regExpSubMatchesItems[objID]; exists {
		return vm.dispatchRegExpSubMatchesPropertyGet(subMatches, member), true
	}

	if subMatchValue, exists := vm.regExpSubMatchValueItems[objID]; exists {
		return vm.dispatchRegExpSubMatchValuePropertyGet(subMatchValue, member), true
	}

	return Value{Type: VTEmpty}, false
}

// dispatchRegExpPropertySet handles RegExp writable properties.
func (vm *VM) dispatchRegExpPropertySet(objID int64, member string, val Value) bool {
	if re, exists := vm.regExpItems[objID]; exists {
		return vm.dispatchRegExpObjectPropertySet(re, member, val)
	}
	return false
}

// --- RegExp Object Implementation ---

func (vm *VM) dispatchRegExpObjectMethod(re *regExpNativeObject, member string, args []Value) Value {
	switch {
	case strings.EqualFold(member, "Execute"):
		if len(args) < 1 {
			return vm.newRegExpMatchesCollectionObject(nil)
		}
		return vm.regExpExecute(re, args[0].String())
	case strings.EqualFold(member, "Test"):
		if len(args) < 1 {
			return NewBool(false)
		}
		return vm.regExpTest(re, args[0].String())
	case strings.EqualFold(member, "Replace"):
		if len(args) < 2 {
			if len(args) == 1 {
				return args[0]
			}
			return NewString("")
		}
		return vm.regExpReplace(re, args[0].String(), args[1].String())
	}
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchRegExpObjectPropertyGet(re *regExpNativeObject, member string) Value {
	switch {
	case strings.EqualFold(member, "Pattern"):
		return NewString(re.pattern)
	case strings.EqualFold(member, "IgnoreCase"):
		return NewBool(re.ignoreCase)
	case strings.EqualFold(member, "Global"):
		return NewBool(re.global)
	case strings.EqualFold(member, "MultiLine"):
		return NewBool(re.multiLine)
	}
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchRegExpObjectPropertySet(re *regExpNativeObject, member string, val Value) bool {
	switch {
	case strings.EqualFold(member, "Pattern"):
		re.pattern = val.String()
		vm.compileRegExp(re)
		return true
	case strings.EqualFold(member, "IgnoreCase"):
		re.ignoreCase = vm.asBool(val)
		vm.compileRegExp(re)
		return true
	case strings.EqualFold(member, "Global"):
		re.global = vm.asBool(val)
		return true
	case strings.EqualFold(member, "MultiLine"):
		re.multiLine = vm.asBool(val)
		vm.compileRegExp(re)
		return true
	}
	return false
}

func (vm *VM) compileRegExp(re *regExpNativeObject) {
	if re.pattern == "" {
		re.compiled = nil
		re.lastError = ""
		return
	}

	pattern := re.pattern
	if re.multiLine {
		pattern = "(?m)" + pattern
	}
	if re.ignoreCase {
		pattern = "(?i)" + pattern
	}

	compiled, err := regexp.Compile(pattern)
	if err != nil {
		re.compiled = nil
		re.lastError = err.Error()
		vm.raise(vbscript.RegularExpressionSyntaxError, re.lastError)
		return
	}
	re.compiled = compiled
	re.lastError = ""
}

func (vm *VM) regExpExecute(re *regExpNativeObject, input string) Value {
	if re.compiled == nil {
		vm.compileRegExp(re)
		if re.compiled == nil {
			return vm.newRegExpMatchesCollectionObject(nil)
		}
	}

	var matches []*regExpMatch
	if re.global {
		allIndices := re.compiled.FindAllStringSubmatchIndex(input, -1)
		for _, idx := range allIndices {
			matches = append(matches, vm.createRegExpMatch(input, idx))
		}
	} else {
		idx := re.compiled.FindStringSubmatchIndex(input)
		if idx != nil {
			matches = append(matches, vm.createRegExpMatch(input, idx))
		}
	}

	return vm.newRegExpMatchesCollectionObject(matches)
}

func (vm *VM) createRegExpMatch(input string, idx []int) *regExpMatch {
	matchValue := input[idx[0]:idx[1]]

	subValues := make([]string, 0)
	// Submatches start from index 2 (index 0 and 1 are the whole match)
	for i := 2; i < len(idx); i += 2 {
		if idx[i] >= 0 && idx[i+1] >= 0 {
			subValues = append(subValues, input[idx[i]:idx[i+1]])
		} else {
			subValues = append(subValues, "")
		}
	}

	return &regExpMatch{
		value:  matchValue,
		index:  idx[0],
		length: len(matchValue),
		subMatches: &regExpSubMatches{
			values: subValues,
		},
	}
}
func (vm *VM) regExpTest(re *regExpNativeObject, input string) Value {
	if re.compiled == nil {
		vm.compileRegExp(re)
		if re.compiled == nil {
			return NewBool(false)
		}
	}
	return NewBool(re.compiled.MatchString(input))
}

func (vm *VM) regExpReplace(re *regExpNativeObject, input string, replacement string) Value {
	if re.compiled == nil {
		vm.compileRegExp(re)
		if re.compiled == nil {
			return NewString(input)
		}
	}

	if !re.global {
		// Replace first occurrence
		idx := re.compiled.FindStringSubmatchIndex(input)
		if idx == nil {
			return NewString(input)
		}
		res := input[:idx[0]] + vm.processRegExpReplacement(replacement, input, idx) + input[idx[1]:]
		return NewString(res)
	}

	// Global replace
	var sb strings.Builder
	lastIdx := 0
	allIndices := re.compiled.FindAllStringSubmatchIndex(input, -1)
	for _, idx := range allIndices {
		sb.WriteString(input[lastIdx:idx[0]])
		sb.WriteString(vm.processRegExpReplacement(replacement, input, idx))
		lastIdx = idx[1]
	}
	sb.WriteString(input[lastIdx:])
	return NewString(sb.String())
}

func (vm *VM) processRegExpReplacement(replacement string, input string, idx []int) string {
	// VBScript supports:
	// $& or $0 - entire match
	// $1, $2... - capture groups
	// $` - text before match
	// $' - text after match

	res := replacement
	matchText := input[idx[0]:idx[1]]

	res = strings.ReplaceAll(res, "$&", matchText)
	res = strings.ReplaceAll(res, "$0", matchText)

	if strings.Contains(res, "$` ") || strings.Contains(res, "$`") {
		res = strings.ReplaceAll(res, "$`", input[:idx[0]])
	}
	if strings.Contains(res, "$' ") || strings.Contains(res, "$'") {
		res = strings.ReplaceAll(res, "$'", input[idx[1]:])
	}

	// Handle $1, $2...
	for i := 1; i <= (len(idx)/2)-1; i++ {
		placeholder := fmt.Sprintf("$%d", i)
		if strings.Contains(res, placeholder) {
			groupText := ""
			if idx[i*2] >= 0 && idx[i*2+1] >= 0 {
				groupText = input[idx[i*2]:idx[i*2+1]]
			}
			res = strings.ReplaceAll(res, placeholder, groupText)
		}
	}

	return res
}

// --- Matches Collection Implementation ---

func (vm *VM) newRegExpMatchesCollectionObject(matches []*regExpMatch) Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.regExpMatchesCollectionItems[objID] = &regExpMatchesCollection{
		matches: matches,
	}
	return Value{Type: VTNativeObject, Num: objID}
}

func (vm *VM) dispatchRegExpMatchesCollectionMethod(m *regExpMatchesCollection, member string, args []Value) Value {
	switch {
	case member == "" || strings.EqualFold(member, "Item"):
		if len(args) > 0 {
			idx := vm.asInt(args[0])
			if idx >= 0 && idx < len(m.matches) {
				return vm.newRegExpMatchObject(m.matches[idx])
			}
		}
	case strings.EqualFold(member, "Count"):
		return NewInteger(int64(len(m.matches)))
	}
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchRegExpMatchesCollectionPropertyGet(m *regExpMatchesCollection, member string) Value {
	switch {
	case strings.EqualFold(member, "Count") || strings.EqualFold(member, "Length"):
		return NewInteger(int64(len(m.matches)))
	}
	return Value{Type: VTEmpty}
}

// --- Match Object Implementation ---

func (vm *VM) newRegExpMatchObject(match *regExpMatch) Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.regExpMatchItems[objID] = match
	return Value{Type: VTNativeObject, Num: objID}
}

func (vm *VM) dispatchRegExpMatchMethod(m *regExpMatch, member string, args []Value) Value {
	switch {
	case strings.EqualFold(member, "SubMatches"):
		if len(args) > 0 {
			// rs.SubMatches(0)
			idx := vm.asInt(args[0])
			if idx >= 0 && idx < len(m.subMatches.values) {
				return NewString(m.subMatches.values[idx])
			}
			return Value{Type: VTEmpty}
		}
		return vm.newRegExpSubMatchesObject(m.subMatches)
	}
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchRegExpMatchPropertyGet(m *regExpMatch, member string) Value {
	switch {
	case strings.EqualFold(member, "Value") || member == "":
		return NewString(m.value)
	case strings.EqualFold(member, "FirstIndex"):
		return NewInteger(int64(m.index))
	case strings.EqualFold(member, "Length"):
		return NewInteger(int64(m.length))
	case strings.EqualFold(member, "SubMatches"):
		return vm.newRegExpSubMatchesObject(m.subMatches)
	}
	return Value{Type: VTEmpty}
}

// --- SubMatches Implementation ---

func (vm *VM) newRegExpSubMatchesObject(sm *regExpSubMatches) Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.regExpSubMatchesItems[objID] = sm
	return Value{Type: VTNativeObject, Num: objID}
}

func (vm *VM) dispatchRegExpSubMatchesMethod(sm *regExpSubMatches, member string, args []Value) Value {
	switch {
	case member == "" || strings.EqualFold(member, "Item"):
		if len(args) > 0 {
			idx := vm.asInt(args[0])
			if idx >= 0 && idx < len(sm.values) {
				return vm.newRegExpSubMatchValueObject(sm.values[idx])
			}
		}
	case strings.EqualFold(member, "Count"):
		return NewInteger(int64(len(sm.values)))
	}
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchRegExpSubMatchesPropertyGet(sm *regExpSubMatches, member string) Value {
	switch {
	case strings.EqualFold(member, "Count"):
		return NewInteger(int64(len(sm.values)))
	default:
		// Default property for SubMatches is Item.
		var idx int
		if _, err := fmt.Sscanf(member, "%d", &idx); err == nil {
			if idx >= 0 && idx < len(sm.values) {
				return vm.newRegExpSubMatchValueObject(sm.values[idx])
			}
		}
	}
	return Value{Type: VTEmpty}
}

// newRegExpSubMatchValueObject allocates one SubMatch value object.
func (vm *VM) newRegExpSubMatchValueObject(value string) Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.regExpSubMatchValueItems[objID] = &regExpSubMatchValue{value: value}
	return Value{Type: VTNativeObject, Num: objID}
}

// dispatchRegExpSubMatchValueMethod resolves callable members on one SubMatch value object.
func (vm *VM) dispatchRegExpSubMatchValueMethod(sm *regExpSubMatchValue, member string, args []Value) Value {
	_ = args
	switch {
	case member == "" || strings.EqualFold(member, "Value"):
		return NewString(sm.value)
	}
	return Value{Type: VTEmpty}
}

// dispatchRegExpSubMatchValuePropertyGet resolves readable properties on one SubMatch value object.
func (vm *VM) dispatchRegExpSubMatchValuePropertyGet(sm *regExpSubMatchValue, member string) Value {
	switch {
	case member == "" || strings.EqualFold(member, "Value"):
		return NewString(sm.value)
	case strings.EqualFold(member, "Length"):
		return NewInteger(int64(len(sm.value)))
	}
	return Value{Type: VTEmpty}
}

// --- VM Enumeration Helper ---

func (vm *VM) regExpMatchesToValues(objID int64) []Value {
	m, exists := vm.regExpMatchesCollectionItems[objID]
	if !exists {
		return nil
	}
	res := make([]Value, len(m.matches))
	for i, match := range m.matches {
		res[i] = vm.newRegExpMatchObject(match)
	}
	return res
}
