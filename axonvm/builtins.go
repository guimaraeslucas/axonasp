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
	"math"
	"math/rand"
	"net/url"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"time"
	"unicode"

	"g3pix.com.br/axonasp/vbscript"
)

// BuiltinFunc is the signature for all VBScript built-in functions.
type BuiltinFunc func(vm *VM, args []Value) (Value, error)

// builtinVBRuntimeError maps one builtin failure to a concrete VBScript runtime error code.
type builtinVBRuntimeError struct {
	code    vbscript.VBSyntaxErrorCode
	message string
}

func (e builtinVBRuntimeError) Error() string {
	if e.message != "" {
		return e.message
	}
	return e.code.String()
}

// newBuiltinVBRuntimeError builds one runtime error compatible with VBScript Err.Number semantics.
func newBuiltinVBRuntimeError(code vbscript.VBSyntaxErrorCode, message string) error {
	return builtinVBRuntimeError{code: code, message: message}
}

// bindBuiltin adapts argument-only builtins to VM-aware builtins.
func bindBuiltin(fn func(args []Value) (Value, error)) BuiltinFunc {
	return func(_ *VM, args []Value) (Value, error) {
		return fn(args)
	}
}

var BuiltinRegistry []BuiltinFunc
var BuiltinNames []string
var BuiltinIndex = make(map[string]int)
var globalAxonFunctionsInitialized bool

// lastRandomValue stores the last value returned by Rnd() for Rnd(0) calls.
var lastRandomValue float64 = 0

// RegisterBuiltin adds a Go function to the VBScript global scope.
func RegisterBuiltin(name string, fn BuiltinFunc) {
	idx := len(BuiltinRegistry)
	BuiltinRegistry = append(BuiltinRegistry, fn)
	BuiltinNames = append(BuiltinNames, name)
	BuiltinIndex[strings.ToLower(name)] = idx
}

// InitGlobalAxonFunctions conditionally injects Axon custom functions into the global VBScript built-in registry.
// This must be called at process startup before any compiler or VM instance is created.
func InitGlobalAxonFunctions(enable bool) {
	if !enable || globalAxonFunctionsInitialized {
		return
	}

	for i, name := range AxonGlobalFunctionNames {
		if _, exists := GetBuiltinIndex(name); exists {
			continue
		}
		if i >= len(AxonGlobalFunctionPointers) {
			break
		}
		RegisterBuiltin(name, AxonGlobalFunctionPointers[i])
	}

	globalAxonFunctionsInitialized = true
}

// --- Type Checking and Variants ---

func vbsIsEmpty(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewBool(true), nil
	}
	return NewBool(args[0].Type == VTEmpty), nil
}

func vbsIsNull(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewBool(false), nil
	}
	return NewBool(args[0].Type == VTNull), nil
}

func vbsIsObject(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewBool(false), nil
	}
	return NewBool(args[0].Type == VTObject || args[0].Type == VTNativeObject), nil
}

func vbsErlVM(vm *VM, args []Value) (Value, error) {
	if vm == nil || vm.errObject == nil {
		return NewInteger(0), nil
	}
	return NewInteger(int64(vm.errObject.Line)), nil
}

func vbsTypeName(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString("Empty"), nil
	}
	switch args[0].Type {
	case VTEmpty:
		return NewString("Empty"), nil
	case VTNull:
		return NewString("Null"), nil
	case VTBool:
		return NewString("Boolean"), nil
	case VTInteger:
		return NewString("Integer"), nil
	case VTDouble:
		return NewString("Double"), nil
	case VTString:
		return NewString("String"), nil
	case VTDate:
		return NewString("Date"), nil
	case VTArray:
		return NewString("Variant()"), nil
	case VTNativeObject, VTObject:
		return NewString("Object"), nil
	default:
		return NewString("Unknown"), nil
	}
}

// resolveNativeTypeName returns the Classic ASP-compatible TypeName string for one native object ID.
// TypeName must match the values that Classic ASP / WinScript returns so that ASP library code
// (e.g. the asplite.asp JSON serializer) can branch correctly on the object type.
func (vm *VM) resolveNativeTypeName(objID int64) string {
	if vm == nil {
		return ""
	}
	// ADODB objects — Classic ASP returns the short COM class name, not "Object".
	if _, ok := vm.adodbRecordsetItems[objID]; ok {
		return "Recordset"
	}
	if _, ok := vm.adodbConnectionItems[objID]; ok {
		return "Connection"
	}
	if _, ok := vm.adodbCommandItems[objID]; ok {
		return "Command"
	}
	if _, ok := vm.adodbParameterItems[objID]; ok {
		return "Parameter"
	}
	if _, ok := vm.adodbFieldItems[objID]; ok {
		return "Field"
	}
	if _, ok := vm.adodbFieldsCollectionItems[objID]; ok {
		return "Fields"
	}
	if _, ok := vm.adodbParametersCollectionItems[objID]; ok {
		return "Parameters"
	}
	if _, ok := vm.adodbErrorItems[objID]; ok {
		return "Error"
	}
	if _, ok := vm.adodbErrorsCollectionItems[objID]; ok {
		return "Errors"
	}
	if _, ok := vm.adodbStreamItems[objID]; ok {
		return "Stream"
	}
	if _, ok := vm.dictionaryItems[objID]; ok {
		return "Dictionary"
	}
	return ""
}

// vbsTypeNameVM resolves TypeName with VM context, including runtime class instance names.
func vbsTypeNameVM(vm *VM, args []Value) (Value, error) {
	if len(args) >= 1 && vm != nil {
		switch args[0].Type {
		case VTObject:
			if instance, exists := vm.runtimeClassItems[args[0].Num]; exists && instance != nil {
				trimmedClassName := strings.TrimSpace(instance.ClassName)
				if trimmedClassName != "" {
					return NewString(trimmedClassName), nil
				}
			}
		case VTNativeObject:
			typeName := vm.resolveNativeTypeName(args[0].Num)
			if typeName != "" {
				return NewString(typeName), nil
			}
		}
	}
	return vbsTypeName(args)
}

// vbsArray creates a VBScript Variant array from function arguments.
func vbsArray(args []Value) (Value, error) {
	values := make([]Value, len(args))
	copy(values, args)
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}, nil
}

// vbsIsArray checks whether a value is an initialized VBScript array.
func vbsIsArray(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewBool(false), nil
	}
	return NewBool(args[0].Type == VTArray && args[0].Arr != nil), nil
}

// vbsLBound returns the lower bound for a VBScript array dimension.
func vbsLBound(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}

	dimension := 1
	if len(args) >= 2 {
		dimension = int(args[1].Num)
	}

	lower, _, ok := arrayBounds(args[0], dimension)
	if !ok {
		return NewInteger(0), nil
	}

	return NewInteger(int64(lower)), nil
}

// vbsUBound returns the upper bound for a VBScript array dimension.
func vbsUBound(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(-1), nil
	}

	dimension := 1
	if len(args) >= 2 {
		dimension = int(args[1].Num)
	}

	_, upper, ok := arrayBounds(args[0], dimension)
	if !ok {
		return NewInteger(-1), nil
	}

	return NewInteger(int64(upper)), nil
}

// vbsLBoundVM applies VBScript runtime error semantics for invalid LBound calls.
func vbsLBoundVM(vm *VM, args []Value) (Value, error) {
	result, _ := vbsLBound(args)
	if len(args) < 1 {
		vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
		return result, nil
	}

	dimension := 1
	if len(args) >= 2 {
		dimension = int(args[1].Num)
	}
	if _, _, ok := arrayBounds(args[0], dimension); !ok {
		vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
	}

	return result, nil
}

// vbsUBoundVM applies VBScript runtime error semantics for invalid UBound calls.
func vbsUBoundVM(vm *VM, args []Value) (Value, error) {
	result, _ := vbsUBound(args)
	if len(args) < 1 {
		vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
		return result, nil
	}

	dimension := 1
	if len(args) >= 2 {
		dimension = int(args[1].Num)
	}
	if _, _, ok := arrayBounds(args[0], dimension); !ok {
		vm.raise(vbscript.SubscriptOutOfRange, "Subscript out of range")
	}

	return result, nil
}

// --- String and Binary Manipulation ---

func VbsLen(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(len([]rune(args[0].String())))), nil
}

func vbsLenB(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(len(ansiBytes(args[0].String())))), nil
}

func VbsUCase(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	return NewString(strings.ToUpper(args[0].String())), nil
}

func vbsLCase(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	return NewString(strings.ToLower(args[0].String())), nil
}

func vbsTrim(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	return NewString(strings.TrimSpace(args[0].String())), nil
}

func vbsMid(args []Value) (Value, error) {
	if len(args) < 2 {
		return Value{Type: VTEmpty}, nil
	}
	runes := []rune(args[0].String())
	start := int(args[1].Num) - 1
	if start < 0 {
		start = 0
	}
	if start >= len(runes) {
		return NewString(""), nil
	}

	if len(args) >= 3 {
		length := int(args[2].Num)
		if start+length > len(runes) {
			length = len(runes) - start
		}
		return NewString(string(runes[start : start+length])), nil
	}
	return NewString(string(runes[start:])), nil
}

func vbsLeft(args []Value) (Value, error) {
	if len(args) < 2 {
		return Value{Type: VTEmpty}, nil
	}
	runes := []rune(args[0].String())
	length := int(args[1].Num)
	if length <= 0 {
		return NewString(""), nil
	}
	if length > len(runes) {
		length = len(runes)
	}
	return NewString(string(runes[:length])), nil
}

func vbsRight(args []Value) (Value, error) {
	if len(args) < 2 {
		return Value{Type: VTEmpty}, nil
	}
	runes := []rune(args[0].String())
	length := int(args[1].Num)
	if length <= 0 {
		return NewString(""), nil
	}
	if length > len(runes) {
		length = len(runes)
	}
	return NewString(string(runes[len(runes)-length:])), nil
}

func vbsAsc(args []Value) (Value, error) {
	if len(args) < 1 || args[0].String() == "" {
		return NewInteger(0), nil
	}
	return NewInteger(int64(args[0].String()[0])), nil
}

func vbsChr(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	return NewString(string(byte(args[0].Num))), nil
}

func vbsInStrVM(vm *VM, args []Value) (Value, error) {
	// InStr([start,] string1, string2[, compare])
	if len(args) < 2 {
		return NewInteger(0), nil
	}

	start := 1
	s1 := ""
	s2 := ""
	compare := 0

	isNumericType := func(v Value) bool {
		switch v.Type {
		case VTBool, VTInteger, VTDouble, VTEmpty:
			return true
		default:
			return false
		}
	}

	switch len(args) {
	case 2:
		s1 = args[0].String()
		s2 = args[1].String()
	case 3:
		if isNumericType(args[0]) {
			start = int(args[0].Num)
			s1 = args[1].String()
			s2 = args[2].String()
		} else {
			s1 = args[0].String()
			s2 = args[1].String()
			compare = int(args[2].Num)
		}
	default:
		start = int(args[0].Num)
		s1 = args[1].String()
		s2 = args[2].String()
		compare = int(args[3].Num)
	}

	if start <= 0 {
		return NewEmpty(), newBuiltinVBRuntimeError(vbscript.InvalidProcedureCallOrArgument, vbscript.InvalidProcedureCallOrArgument.String())
	}

	runes1 := []rune(s1)
	runes2 := []rune(s2)
	if s2 == "" {
		if start > len(runes1)+1 {
			return NewInteger(0), nil
		}
		return NewInteger(int64(start)), nil
	}

	if start > len(runes1) || len(runes2) == 0 {
		return NewInteger(0), nil
	}

	textCompare := compare == 1 || (compare == -1 && vm != nil && vm.optionCompare == 1)
	for idx := start - 1; idx+len(runes2) <= len(runes1); idx++ {
		segment := string(runes1[idx : idx+len(runes2)])
		if (textCompare && strings.EqualFold(segment, s2)) || (!textCompare && segment == s2) {
			return NewInteger(int64(idx + 1)), nil
		}
	}

	return NewInteger(0), nil
}

// vbsReplace implements VBScript Replace(expression, find, replace[, start[, count[, compare]]]).
func vbsReplaceVM(vm *VM, args []Value) (Value, error) {
	if len(args) < 3 {
		return NewString(""), nil
	}
	if args[0].Type == VTNull {
		return NewNull(), nil
	}

	expression := args[0].String()
	find := args[1].String()
	replacement := args[2].String()

	start := 1
	if len(args) >= 4 {
		start = int(args[3].Num)
	}
	if start <= 0 {
		return NewEmpty(), newBuiltinVBRuntimeError(vbscript.InvalidProcedureCallOrArgument, vbscript.InvalidProcedureCallOrArgument.String())
	}

	count := -1
	if len(args) >= 5 {
		count = int(args[4].Num)
	}

	compare := 0
	if len(args) >= 6 {
		compare = int(args[5].Num)
	}

	runes := []rune(expression)
	if start > len(runes) {
		return NewString(""), nil
	}

	targetRunes := runes[start-1:]
	if len(targetRunes) == 0 {
		return NewString(""), nil
	}

	if find == "" {
		return NewString(string(targetRunes)), nil
	}

	if count == 0 {
		return NewString(string(targetRunes)), nil
	}

	findRunes := []rune(find)
	if len(findRunes) == 0 {
		return NewString(string(targetRunes)), nil
	}

	textCompare := compare == 1 || (compare == -1 && vm != nil && vm.optionCompare == 1)
	sourceRunes := targetRunes
	patternRunes := findRunes
	if textCompare {
		sourceRunes = make([]rune, len(targetRunes))
		for i, r := range targetRunes {
			sourceRunes[i] = unicode.ToLower(r)
		}
		patternRunes = make([]rune, len(findRunes))
		for i, r := range findRunes {
			patternRunes[i] = unicode.ToLower(r)
		}
	}

	// Use one KMP scan over runes to avoid repeated []rune/string churn inside the hot loop.
	prefix := buildKMPPrefix(patternRunes)
	var builder strings.Builder
	builder.Grow(len(expression) + len(replacement))

	patLen := len(patternRunes)
	lastEmit := 0
	replaced := 0
	j := 0
	for i := 0; i < len(sourceRunes); i++ {
		for j > 0 && sourceRunes[i] != patternRunes[j] {
			j = prefix[j-1]
		}
		if sourceRunes[i] == patternRunes[j] {
			j++
		}

		if j == patLen {
			matchStart := i - patLen + 1
			writeRuneSlice(&builder, targetRunes[lastEmit:matchStart])
			builder.WriteString(replacement)
			replaced++
			lastEmit = matchStart + patLen
			j = 0
			if count >= 0 && replaced >= count {
				break
			}
		}
	}

	writeRuneSlice(&builder, targetRunes[lastEmit:])
	return NewString(builder.String()), nil
}

// buildKMPPrefix computes the failure table used by KMP search over rune slices.
func buildKMPPrefix(pattern []rune) []int {
	prefix := make([]int, len(pattern))
	j := 0
	for i := 1; i < len(pattern); i++ {
		for j > 0 && pattern[i] != pattern[j] {
			j = prefix[j-1]
		}
		if pattern[i] == pattern[j] {
			j++
			prefix[i] = j
		}
	}
	return prefix
}

// writeRuneSlice appends rune content to one strings.Builder without intermediate string allocation.
func writeRuneSlice(builder *strings.Builder, runes []rune) {
	for _, r := range runes {
		builder.WriteRune(r)
	}
}

// vbsMsgBox always raises one runtime error in ASP because interactive desktop UI is unsupported.
func vbsMsgBox(args []Value) (Value, error) {
	_ = args
	return NewEmpty(), newBuiltinVBRuntimeError(vbscript.InvalidProcedureCallOrArgument, fmt.Sprintf("%d: %s (MsgBox)", ErrInteractiveFunctionNotSupportedInASP, ErrInteractiveFunctionNotSupportedInASP.String()))
}

// vbsInputBox always raises one runtime error in ASP because interactive desktop UI is unsupported.
func vbsInputBox(args []Value) (Value, error) {
	_ = args
	return NewEmpty(), newBuiltinVBRuntimeError(vbscript.InvalidProcedureCallOrArgument, fmt.Sprintf("%d: %s (InputBox)", ErrInteractiveFunctionNotSupportedInASP, ErrInteractiveFunctionNotSupportedInASP.String()))
}

// --- Math and Conversion ---

func vbsAbs(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	if args[0].Type == VTDouble {
		return NewDouble(math.Abs(args[0].Flt)), nil
	}
	val := args[0].Num
	if val < 0 {
		val = -val
	}
	return NewInteger(val), nil
}

func vbsSqr(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewDouble(0), nil
	}
	var f float64
	if args[0].Type == VTDouble {
		f = args[0].Flt
	} else {
		f = float64(args[0].Num)
	}
	return NewDouble(math.Sqrt(f)), nil
}

func vbsInt(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	var f float64
	if args[0].Type == VTDouble {
		f = args[0].Flt
	} else {
		f = float64(args[0].Num)
	}
	return NewInteger(int64(math.Floor(f))), nil
}

func vbsFix(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	var f float64
	if args[0].Type == VTDouble {
		f = args[0].Flt
	} else {
		f = float64(args[0].Num)
	}
	return NewInteger(int64(math.Trunc(f))), nil
}

// vbsRnd implements VBScript Rnd() with full argument support.
// Rnd()    - returns random number between 0 and 1
// Rnd(n>0) - returns new random number (n is ignored)
// Rnd(0)   - returns last random number generated
// Rnd(n<0) - reseeds with -n and returns new random number
func vbsRnd(args []Value) (Value, error) {
	if len(args) == 0 {
		// No argument: return new random value
		lastRandomValue = rand.Float64()
		return NewDouble(lastRandomValue), nil
	}

	// Convert argument to number
	arg := args[0]
	var seed float64

	switch arg.Type {
	case VTInteger:
		seed = float64(arg.Num)
	case VTDouble:
		seed = arg.Flt
	case VTBool:
		if arg.Num != 0 {
			seed = 1
		} else {
			seed = 0
		}
	case VTEmpty:
		seed = 0
	default:
		seed = 0
	}

	if seed < 0 {
		// Reseed with absolute value
		rand.Seed(int64(-seed))
		lastRandomValue = rand.Float64()
		return NewDouble(lastRandomValue), nil
	} else if seed == 0 {
		// Return last random number
		return NewDouble(lastRandomValue), nil
	} else {
		// seed > 0: return new random value (seed parameter is ignored in classic ASP)
		lastRandomValue = rand.Float64()
		return NewDouble(lastRandomValue), nil
	}
}

func vbsRandomize(args []Value) (Value, error) {
	seed := time.Now().UnixNano()
	if len(args) >= 1 {
		seed = args[0].Num
	}
	rand.Seed(seed)
	return Value{Type: VTEmpty}, nil
}

// vbsParseNumericString parses VBScript numeric text, including &H and &O string forms.
func vbsParseNumericString(text string) (float64, vbscript.VBSyntaxErrorCode, bool) {
	trimmed := strings.TrimSpace(text)
	if trimmed == "" {
		return 0, vbscript.TypeMismatch, false
	}

	sign := 1.0
	if strings.HasPrefix(trimmed, "+") {
		trimmed = strings.TrimSpace(trimmed[1:])
	} else if strings.HasPrefix(trimmed, "-") {
		sign = -1
		trimmed = strings.TrimSpace(trimmed[1:])
	}

	if trimmed == "" {
		return 0, vbscript.TypeMismatch, false
	}

	lower := strings.ToLower(trimmed)
	if strings.HasPrefix(lower, "&h") {
		digits := trimmed[2:]
		if digits == "" {
			return 0, vbscript.TypeMismatch, false
		}
		parsed, err := strconv.ParseUint(digits, 16, 64)
		if err != nil {
			return 0, vbscript.TypeMismatch, false
		}
		return sign * float64(parsed), 0, true
	}

	if strings.HasPrefix(lower, "&o") {
		digits := trimmed[2:]
		if digits == "" {
			return 0, vbscript.TypeMismatch, false
		}
		parsed, err := strconv.ParseUint(digits, 8, 64)
		if err != nil {
			return 0, vbscript.TypeMismatch, false
		}
		return sign * float64(parsed), 0, true
	}

	parsed, err := strconv.ParseFloat(trimmed, 64)
	if err != nil || math.IsNaN(parsed) || math.IsInf(parsed, 0) {
		return 0, vbscript.TypeMismatch, false
	}

	return sign * parsed, 0, true
}

// vbsCoerceNumericValue converts a VM value to a numeric value using VBScript coercion rules.
func vbsCoerceNumericValue(value Value) (float64, vbscript.VBSyntaxErrorCode, bool) {
	switch value.Type {
	case VTEmpty:
		return 0, 0, true
	case VTNull:
		return 0, vbscript.InvalidUseOfNull, false
	case VTBool, VTInteger:
		return float64(value.Num), 0, true
	case VTDouble:
		return value.Flt, 0, true
	case VTDate:
		return float64(value.Num), 0, true
	case VTString:
		return vbsParseNumericString(value.Str)
	default:
		return 0, vbscript.TypeMismatch, false
	}
}

// vbsConvertRoundedIntegerVM converts one value to an integer range using VBScript banker's rounding.
func vbsConvertRoundedIntegerVM(vm *VM, args []Value, minValue int64, maxValue int64) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}

	numericValue, code, ok := vbsCoerceNumericValue(args[0])
	if !ok {
		if vm != nil {
			vm.raise(code, code.String())
		}
		return NewInteger(0), nil
	}

	rounded := math.RoundToEven(numericValue)
	if rounded < float64(minValue) || rounded > float64(maxValue) {
		if vm != nil {
			vm.raise(vbscript.Overflow, vbscript.Overflow.String())
		}
		return NewInteger(0), nil
	}

	return NewInteger(int64(rounded)), nil
}

// vbsCIntVM converts values using VBScript CInt semantics, including banker's rounding.
func vbsCIntVM(vm *VM, args []Value) (Value, error) {
	return vbsConvertRoundedIntegerVM(vm, args, -32768, 32767)
}

// vbsCDblVM converts values using VBScript CDbl semantics and VM error propagation.
func vbsCDblVM(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewDouble(0), nil
	}

	numericValue, code, ok := vbsCoerceNumericValue(args[0])
	if !ok {
		if vm != nil {
			vm.raise(code, code.String())
		}
		return NewDouble(0), nil
	}

	return NewDouble(numericValue), nil
}

// vbsCStrVM converts a value to a string, using VM for locale-aware date formatting.
func vbsCStrVM(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	// Use vm.valueToString for proper locale-aware formatting
	return NewString(vm.valueToString(args[0])), nil
}

// --- Formatting and Locality ---

func vbsRGB(args []Value) (Value, error) {
	if len(args) < 3 {
		return NewInteger(0), nil
	}
	r := args[0].Num & 0xFF
	g := args[1].Num & 0xFF
	b := args[2].Num & 0xFF
	return NewInteger(r + (g * 256) + (b * 65536)), nil
}

func vbsEscape(args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	return NewString(url.QueryEscape(args[0].String())), nil
}

// vbsIIf implements the VBScript IIf(expr, truePart, falsePart) function.
// Both truePart and falsePart are always evaluated before the call.
func vbsIIf(args []Value) (Value, error) {
	if len(args) < 3 {
		return NewEmpty(), nil
	}
	cond := args[0]
	switch cond.Type {
	case VTBool:
		if cond.Num != 0 {
			return args[1], nil
		}
	case VTInteger:
		if cond.Num != 0 {
			return args[1], nil
		}
	case VTDouble:
		if cond.Flt != 0 {
			return args[1], nil
		}
	case VTString:
		if cond.Str != "" && cond.Str != "0" {
			return args[1], nil
		}
	case VTNull, VTEmpty:
		// Null and Empty are falsy
	default:
		return args[1], nil
	}
	return args[2], nil
}

// toArrayBound converts a VM value into an upper-bound integer for Dim and ReDim operations.
func toArrayBound(value Value) (int, error) {
	switch value.Type {
	case VTInteger, VTBool:
		return int(value.Num), nil
	case VTDouble:
		return int(value.Flt), nil
	case VTString:
		parsed, err := strconv.Atoi(strings.TrimSpace(value.Str))
		if err != nil {
			return 0, fmt.Errorf("invalid array bound: %s", value.Str)
		}
		return parsed, nil
	case VTEmpty:
		return 0, nil
	default:
		return 0, fmt.Errorf("invalid array bound type: %s", value.String())
	}
}

// buildDimArray creates a zero-based VBScript array tree for one or more dimensions.
func buildDimArray(bounds []int) *VBArray {
	if len(bounds) == 0 {
		return NewVBArray(0, 0)
	}

	size := bounds[0] + 1
	if size < 0 {
		size = 0
	}

	array := NewVBArray(0, size)
	if len(bounds) == 1 {
		return array
	}

	for index := range array.Values {
		array.Values[index] = ValueFromVBArray(buildDimArray(bounds[1:]))
	}

	return array
}

// vbArrayUpperBounds returns one upper-bound list from a VBArray shape.
func vbArrayUpperBounds(array *VBArray) []int {
	if array == nil {
		return nil
	}

	bounds := make([]int, 0, 4)
	current := array
	for {
		bounds = append(bounds, len(current.Values)-1)
		if len(current.Values) == 0 {
			break
		}
		next, ok := toVBArray(current.Values[0])
		if !ok {
			break
		}
		current = next
	}
	return bounds
}

// copyPreservedArray copies common array elements into a resized array while preserving VBScript shape rules.
func copyPreservedArray(target *VBArray, source *VBArray, remainingBounds []int) {
	if target == nil || source == nil {
		return
	}

	limit := len(target.Values)
	if len(source.Values) < limit {
		limit = len(source.Values)
	}

	if len(remainingBounds) == 1 {
		copy(target.Values[:limit], source.Values[:limit])
		return
	}

	for index := 0; index < limit; index++ {
		sourceChild, ok := toVBArray(source.Values[index])
		if !ok {
			continue
		}
		target.Values[index] = ValueFromVBArray(buildDimArray(remainingBounds[1:]))
		targetChild, _ := toVBArray(target.Values[index])
		copyPreservedArray(targetChild, sourceChild, remainingBounds[1:])
	}
}

// vbsAxonDimArray allocates a zero-based VBScript array for Dim declarations.
func vbsAxonDimArray(args []Value) (Value, error) {
	bounds := make([]int, len(args))
	for index, arg := range args {
		bound, err := toArrayBound(arg)
		if err != nil {
			return NewEmpty(), err
		}
		bounds[index] = bound
	}
	return ValueFromVBArray(buildDimArray(bounds)), nil
}

// vbsAxonRedimArray resizes a VBScript array without preserving existing contents.
func vbsAxonRedimArray(args []Value) (Value, error) {
	if len(args) == 0 {
		return ValueFromVBArray(buildDimArray(nil)), nil
	}
	return vbsAxonDimArray(args[1:])
}

// vbsAxonRedimPreserveArray resizes a VBScript array while preserving common contents.
func vbsAxonRedimPreserveArray(args []Value) (Value, error) {
	if len(args) == 0 {
		return ValueFromVBArray(buildDimArray(nil)), nil
	}

	bounds := make([]int, 0, len(args)-1)
	for _, arg := range args[1:] {
		bound, err := toArrayBound(arg)
		if err != nil {
			return NewEmpty(), err
		}
		bounds = append(bounds, bound)
	}

	resized := buildDimArray(bounds)
	if existing, ok := toVBArray(args[0]); ok {
		existingBounds := vbArrayUpperBounds(existing)
		if len(existingBounds) > 1 && len(bounds) == len(existingBounds) {
			for index := 0; index < len(bounds)-1; index++ {
				if bounds[index] != existingBounds[index] {
					return NewEmpty(), newBuiltinVBRuntimeError(vbscript.InvalidProcedureCallOrArgument, vbscript.InvalidProcedureCallOrArgument.String())
				}
			}
		}
		copyPreservedArray(resized, existing, bounds)
	}

	return ValueFromVBArray(resized), nil
}

// vbsAxonEnumValues normalizes supported enumerable values into a zero-based array snapshot.
func vbsAxonEnumValues(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return ValueFromVBArray(NewVBArrayFromValues(0, nil)), nil
	}

	target := args[0]
	if target.Type == VTArray && target.Arr != nil {
		return target, nil
	}

	if target.Type != VTNativeObject || vm == nil {
		return ValueFromVBArray(NewVBArrayFromValues(0, nil)), nil
	}

	if target.Num == nativeObjectSessionContents {
		keys := vm.host.Session().GetAllKeys()
		sort.Strings(keys)
		values := make([]Value, 0, len(keys))
		for _, key := range keys {
			values = append(values, NewString(key))
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	if target.Num == nativeObjectApplicationContents {
		contents := vm.host.Application().GetContentsCopy()
		keys := make([]string, 0, len(contents))
		for key := range contents {
			keys = append(keys, key)
		}
		sort.Strings(keys)
		values := make([]Value, 0, len(keys))
		for _, key := range keys {
			values = append(values, NewString(key))
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	if target.Num == nativeObjectApplicationStaticObjects {
		objects := vm.host.Application().GetStaticObjectsCopy()
		keys := make([]string, 0, len(objects))
		for key, value := range objects {
			if vm.isStaticObjectApplicationValue(value) {
				keys = append(keys, key)
			}
		}
		sort.Strings(keys)
		values := make([]Value, 0, len(keys))
		for _, key := range keys {
			values = append(values, NewString(key))
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	if target.Num == nativeObjectSessionStaticObjects {
		objects := vm.host.Session().GetStaticObjectsCopy()
		keys := make([]string, 0, len(objects))
		for key, value := range objects {
			if vm.isStaticObjectApplicationValue(value) {
				keys = append(keys, key)
			}
		}
		sort.Strings(keys)
		values := make([]Value, 0, len(keys))
		for _, key := range keys {
			values = append(values, NewString(key))
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	// Preserve Scripting.Dictionary insertion order during For Each enumeration.
	if dict, ok := vm.dictionaryItems[target.Num]; ok && dict != nil {
		values := make([]Value, 0, len(dict.keys))
		for i := 0; i < len(dict.keys); i++ {
			values = append(values, dict.keys[i])
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	if values := vm.regExpMatchesToValues(target.Num); values != nil {
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	// RegExp SubMatches collection — enumerate capture groups as plain string values.
	if subMatches, ok := vm.regExpSubMatchesItems[target.Num]; ok && subMatches != nil {
		values := make([]Value, 0, len(subMatches.values))
		for i := 0; i < len(subMatches.values); i++ {
			values = append(values, NewString(subMatches.values[i]))
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	if rs, ok := vm.adodbFieldsCollectionItems[target.Num]; ok && rs != nil {
		values := make([]Value, 0, len(rs.columns))
		for i := 0; i < len(rs.columns); i++ {
			values = append(values, vm.newADODBFieldProxy(rs, rs.columns[i]))
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	// ADOX.Tables collection — yield one VTNativeObject per table item.
	if tables, ok := vm.adoxTablesItems[target.Num]; ok && tables != nil {
		values := make([]Value, 0, len(tables.items))
		for i := 0; i < len(tables.items); i++ {
			values = append(values, vm.newADOXTableObject(tables.items[i]))
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	// ASP intrinsic request/response collections: enumerate as string key arrays.
	if vm.host != nil {
		req := vm.host.Request()
		// enumKeys converts a string key slice to a VTArray of VTString values.
		enumKeys := func(keys []string) (Value, error) {
			values := make([]Value, len(keys))
			for i, k := range keys {
				values[i] = NewString(k)
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
		}
		switch target.Num {
		case nativeRequestQueryString:
			return enumKeys(req.QueryString.GetKeys())
		case nativeRequestForm:
			if req.IsBinaryReadUsed() {
				return enumKeys([]string{})
			}
			req.MarkFormUsed()
			return enumKeys(req.Form.GetKeys())
		case nativeRequestCookies:
			return enumKeys(req.Cookies.GetKeys())
		case nativeRequestServerVariables:
			return enumKeys(req.ServerVars.GetKeys())
		case nativeRequestClientCertificate:
			return enumKeys(req.ClientCertificate.GetKeys())
		case nativeResponseCookies:
			return enumKeys(vm.host.Response().GetCookieKeys())
		}
	}

	// ADODB.Errors collection.
	if conn, ok := vm.adodbErrorsCollectionItems[target.Num]; ok && conn != nil {
		values := make([]Value, 0, len(conn.errors))
		for i := 0; i < len(conn.errors); i++ {
			errObjID := vm.nextDynamicNativeID
			vm.nextDynamicNativeID++
			errCopy := conn.errors[i]
			vm.adodbErrorItems[errObjID] = &errCopy
			values = append(values, Value{Type: VTNativeObject, Num: errObjID})
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	// ADODB.Parameters collection — yield one VTNativeObject per parameter.
	if cmd, ok := vm.adodbParametersCollectionItems[target.Num]; ok && cmd != nil {
		values := make([]Value, 0, len(cmd.parameters))
		for _, p := range cmd.parameters {
			id := vm.nextDynamicNativeID
			vm.nextDynamicNativeID++
			vm.adodbParameterItems[id] = p
			values = append(values, Value{Type: VTNativeObject, Num: id})
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	// MSXML NodeList — yield one VTNativeObject per XML node.
	if nodeList, ok := vm.msxmlNodeListItems[target.Num]; ok && nodeList != nil {
		items := nodeList.Enumeration()
		values := make([]Value, 0, len(items))
		for _, item := range items {
			values = append(values, legacyInterfaceToValue(item, vm))
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	obj, ok := vm.fsoItems[target.Num]
	if !ok || obj == nil {
		return ValueFromVBArray(NewVBArrayFromValues(0, nil)), nil
	}

	if obj.kind == fsoKindDrivesCollection {
		drives := vm.fsoEnumerateDriveNames()
		values := make([]Value, 0, len(drives))
		for i := 0; i < len(drives); i++ {
			values = append(values, vm.newFSODriveObject(drives[i]))
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}

	if obj.kind != fsoKindFilesCollection && obj.kind != fsoKindSubFoldersCollection {
		return ValueFromVBArray(NewVBArrayFromValues(0, nil)), nil
	}

	entries, err := os.ReadDir(obj.path)
	if err != nil {
		return ValueFromVBArray(NewVBArrayFromValues(0, nil)), nil
	}
	sort.Slice(entries, func(i, j int) bool {
		return strings.ToLower(entries[i].Name()) < strings.ToLower(entries[j].Name())
	})

	values := make([]Value, 0, len(entries))
	for _, entry := range entries {
		entryPath := filepath.Join(obj.path, entry.Name())
		if obj.kind == fsoKindFilesCollection {
			if entry.IsDir() {
				continue
			}
			values = append(values, vm.newFSONativeObject(fsoKindFile, entryPath, nil))
			continue
		}

		if !entry.IsDir() {
			continue
		}
		values = append(values, vm.newFSONativeObject(fsoKindFolder, entryPath, nil))
	}

	return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
}

// vbsAxonEnumCount returns the number of elements in a normalized enumerable array.
func vbsAxonEnumCount(args []Value) (Value, error) {
	if len(args) < 1 || args[0].Type != VTArray || args[0].Arr == nil {
		return NewInteger(0), nil
	}
	return NewInteger(int64(len(args[0].Arr.Values))), nil
}

// vbsAxonEnumItem returns one element by zero-based index from a normalized enumerable array.
func vbsAxonEnumItem(args []Value) (Value, error) {
	if len(args) < 2 || args[0].Type != VTArray || args[0].Arr == nil {
		return NewEmpty(), nil
	}

	idx := int(args[1].Num)
	if idx < 0 || idx >= len(args[0].Arr.Values) {
		return NewEmpty(), nil
	}

	return args[0].Arr.Values[idx], nil
}

func init() {
	RegisterBuiltin("__AXON_DIM_ARRAY", bindBuiltin(vbsAxonDimArray))
	RegisterBuiltin("__AXON_REDIM_ARRAY", bindBuiltin(vbsAxonRedimArray))
	RegisterBuiltin("__AXON_REDIM_PRESERVE_ARRAY", bindBuiltin(vbsAxonRedimPreserveArray))
	RegisterBuiltin("__AXON_ENUM_VALUES", vbsAxonEnumValues)
	RegisterBuiltin("__AXON_ENUM_COUNT", bindBuiltin(vbsAxonEnumCount))
	RegisterBuiltin("__AXON_ENUM_ITEM", bindBuiltin(vbsAxonEnumItem))

	// Type Checking
	RegisterBuiltin("IsEmpty", bindBuiltin(vbsIsEmpty))
	RegisterBuiltin("IsNull", bindBuiltin(vbsIsNull))
	RegisterBuiltin("IsArray", bindBuiltin(vbsIsArray))
	RegisterBuiltin("IsObject", bindBuiltin(vbsIsObject))
	RegisterBuiltin("Erl", vbsErlVM)
	RegisterBuiltin("TypeName", vbsTypeNameVM)
	RegisterBuiltin("Array", bindBuiltin(vbsArray))
	RegisterBuiltin("LBound", vbsLBoundVM)
	RegisterBuiltin("UBound", vbsUBoundVM)

	// Strings
	RegisterBuiltin("Len", bindBuiltin(VbsLen))
	RegisterBuiltin("LenB", bindBuiltin(vbsLenB))
	RegisterBuiltin("UCase", bindBuiltin(VbsUCase))
	RegisterBuiltin("LCase", bindBuiltin(vbsLCase))
	RegisterBuiltin("Trim", bindBuiltin(vbsTrim))
	RegisterBuiltin("Mid", bindBuiltin(vbsMid))
	RegisterBuiltin("Left", bindBuiltin(vbsLeft))
	RegisterBuiltin("Right", bindBuiltin(vbsRight))
	RegisterBuiltin("Asc", bindBuiltin(vbsAsc))
	RegisterBuiltin("Chr", bindBuiltin(vbsChr))
	RegisterBuiltin("InStr", vbsInStrVM)
	RegisterBuiltin("Replace", vbsReplaceVM)

	// Math
	RegisterBuiltin("Abs", bindBuiltin(vbsAbs))
	RegisterBuiltin("Sqr", bindBuiltin(vbsSqr))
	RegisterBuiltin("Int", bindBuiltin(vbsInt))
	RegisterBuiltin("Fix", bindBuiltin(vbsFix))
	RegisterBuiltin("Rnd", bindBuiltin(vbsRnd))
	RegisterBuiltin("Randomize", bindBuiltin(vbsRandomize))

	// Conversion
	RegisterBuiltin("CInt", vbsCIntVM)
	RegisterBuiltin("CDbl", vbsCDblVM)
	RegisterBuiltin("CStr", vbsCStrVM)

	// Misc
	RegisterBuiltin("RGB", bindBuiltin(vbsRGB))
	RegisterBuiltin("Escape", bindBuiltin(vbsEscape))
	RegisterBuiltin("IIf", bindBuiltin(vbsIIf))
	RegisterBuiltin("MsgBox", bindBuiltin(vbsMsgBox))
	RegisterBuiltin("InputBox", bindBuiltin(vbsInputBox))
}

// GetBuiltinIndex returns the registry index for a function name.
func GetBuiltinIndex(name string) (int, bool) {
	idx, ok := BuiltinIndex[strings.ToLower(name)]
	return idx, ok
}
