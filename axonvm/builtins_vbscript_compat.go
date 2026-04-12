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
	"bytes"
	"errors"
	"fmt"
	"math"
	"net/url"
	"strconv"
	"strings"
	"sync"
	"time"
	"unicode"
	"unicode/utf16"

	"g3pix.com.br/axonasp/axonconfig"
	"g3pix.com.br/axonasp/vbscript"
)

// builtinDefaults stores fallback global settings loaded from config/axonasp.toml.
type builtinDefaults struct {
	mslcid   int
	timezone string
	charset  string
	location *time.Location
}

var builtinDefaultsOnce sync.Once
var cachedBuiltinDefaults builtinDefaults

// vbsCompatZeroDateSentinelNum returns the internal VTDate payload used when code
// constructs NewDate(time.Time{}). This sentinel must not be treated as a valid
// VBScript date during compatibility checks.
func vbsCompatZeroDateSentinelNum() int64 {
	return NewDate(time.Time{}).Num
}

// vbsCompatCoerceSerialInt coerces one DateSerial/TimeSerial argument to int.
// It mirrors VBScript numeric coercion rules and rejects non-numeric strings.
func vbsCompatCoerceSerialInt(vm *VM, v Value) (int, bool) {
	resolved := resolveCallable(vm, v)
	n, ok := vm.coerceFloatStrict(resolved)
	if !ok {
		return 0, false
	}
	return int(math.RoundToEven(n)), true
}

// loadBuiltinDefaults reads locale-related fallback values from config/axonasp.toml.
func loadBuiltinDefaults() builtinDefaults {
	builtinDefaultsOnce.Do(func() {
		cfg := builtinDefaults{mslcid: int(EnglishUS), timezone: "UTC", charset: "UTF-8", location: time.UTC}
		v := axonconfig.NewViper()
		if m := v.GetInt("global.default_mslcid"); m > 0 {
			cfg.mslcid = m
		}
		if tz := strings.TrimSpace(v.GetString("global.default_timezone")); tz != "" {
			cfg.timezone = tz
		}
		if cs := strings.TrimSpace(v.GetString("global.default_charset")); cs != "" {
			cfg.charset = cs
		}
		if loc, err := ResolveTimezoneLocation(cfg.timezone); err == nil {
			cfg.location = loc
		}
		cachedBuiltinDefaults = cfg
	})
	return cachedBuiltinDefaults
}

// builtinCurrentLCID resolves the effective LCID from session state or config fallback.
func builtinCurrentLCID(vm *VM) int {
	if vm != nil && vm.host != nil && vm.host.SessionEnabled() {
		lcid := vm.host.Session().GetLCID()
		if lcid > 0 {
			return lcid
		}
	}
	return loadBuiltinDefaults().mslcid
}

// builtinCurrentLocation resolves the effective time location for date/time built-ins.
func builtinCurrentLocation(_ *VM) *time.Location {
	cfg := loadBuiltinDefaults()
	if cfg.location != nil {
		return cfg.location
	}
	return time.UTC
}

// builtinLocaleTag returns the Go locale tag for the current LCID.
func builtinLocaleTag(vm *VM) string {
	return builtinLocaleProfileForVM(vm).tag
}

// vbsCompatIsNothing checks whether a value is Nothing/null object equivalent.
func vbsCompatIsNothing(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewBool(true), nil
	}
	if args[0].Type == VTEmpty || args[0].Type == VTNull {
		return NewBool(true), nil
	}
	return NewBool(false), nil
}

// vbsCompatVarType returns a VBScript-compatible VarType code.
func vbsCompatVarType(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	switch args[0].Type {
	case VTEmpty:
		return NewInteger(0), nil
	case VTNull:
		return NewInteger(1), nil
	case VTInteger, VTBool:
		if args[0].Type == VTBool {
			return NewInteger(11), nil
		}
		if args[0].Num >= -32768 && args[0].Num <= 32767 {
			return NewInteger(2), nil
		}
		return NewInteger(3), nil
	case VTDouble:
		return NewInteger(5), nil
	case VTString:
		return NewInteger(8), nil
	case VTDate:
		return NewInteger(7), nil
	case VTArray:
		return NewInteger(8204), nil
	case VTObject, VTNativeObject:
		return NewInteger(9), nil
	default:
		return NewInteger(12), nil
	}
}

// vbsCompatScriptEngine returns the scripting engine name.
func vbsCompatScriptEngine(_ *VM, _ []Value) (Value, error) {
	return NewString("VBScript"), nil
}

// vbsCompatScriptEngineMajorVersion returns the major script engine version.
func vbsCompatScriptEngineMajorVersion(_ *VM, _ []Value) (Value, error) {
	return NewInteger(5), nil
}

// vbsCompatScriptEngineMinorVersion returns the minor script engine version.
func vbsCompatScriptEngineMinorVersion(_ *VM, _ []Value) (Value, error) {
	return NewInteger(8), nil
}

// vbsCompatScriptEngineBuildVersion returns the build script engine version.
func vbsCompatScriptEngineBuildVersion(_ *VM, _ []Value) (Value, error) {
	return NewInteger(16384), nil
}

// vbsCompatLTrim trims leading spaces and tabs.
func vbsCompatLTrim(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	return NewString(strings.TrimLeft(args[0].String(), " \t")), nil
}

// vbsCompatRTrim trims trailing spaces and tabs.
func vbsCompatRTrim(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	return NewString(strings.TrimRight(args[0].String(), " \t")), nil
}

// vbsCompatSpace returns a string with N spaces.
func vbsCompatSpace(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	n := int(args[0].Num)
	if n < 0 {
		n = 0
	}
	return NewString(strings.Repeat(" ", n)), nil
}

// vbsCompatString repeats the first rune from input string.
func vbsCompatString(_ *VM, args []Value) (Value, error) {
	if len(args) < 2 {
		return NewString(""), nil
	}
	n := int(args[0].Num)
	if n <= 0 {
		return NewString(""), nil
	}
	runes := []rune(args[1].String())
	if len(runes) == 0 {
		return NewString(""), nil
	}
	return NewString(strings.Repeat(string(runes[0]), n)), nil
}

// vbsCompatStrReverse reverses a Unicode string.
func vbsCompatStrReverse(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	runes := []rune(args[0].String())
	for i, j := 0, len(runes)-1; i < j; i, j = i+1, j-1 {
		runes[i], runes[j] = runes[j], runes[i]
	}
	return NewString(string(runes)), nil
}

// vbsCompatStrComp compares strings using binary or text compare.
func vbsCompatStrComp(_ *VM, args []Value) (Value, error) {
	if len(args) < 2 {
		return NewInteger(0), nil
	}
	a := args[0].String()
	b := args[1].String()
	compare := 0
	if len(args) > 2 {
		compare = int(args[2].Num)
	}
	if compare == 1 {
		a = strings.ToLower(a)
		b = strings.ToLower(b)
	}
	if a < b {
		return NewInteger(-1), nil
	}
	if a > b {
		return NewInteger(1), nil
	}
	return NewInteger(0), nil
}

// ansiBytes converts a string to byte-oriented ANSI-compatible representation.
func ansiBytes(input string) []byte {
	if input == "" {
		return []byte{}
	}
	buf := make([]byte, 0, len(input))
	for _, r := range input {
		if r <= 0xFF {
			buf = append(buf, byte(r))
		} else {
			buf = append(buf, byte('?'))
		}
	}
	return buf
}

// vbsCompatAscB returns the first byte code from ANSI bytes.
func vbsCompatAscB(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	bs := ansiBytes(args[0].String())
	if len(bs) == 0 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(bs[0])), nil
}

// vbsCompatChrB returns a one-byte ANSI string.
func vbsCompatChrB(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	return NewString(bytesToVBByteString([]byte{byte(args[0].Num & 0xFF)})), nil
}

// vbsCompatAscW returns the codepoint of the first rune.
func vbsCompatAscW(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	r := []rune(args[0].String())
	if len(r) == 0 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(r[0])), nil
}

// vbsCompatChrW returns a string from a Unicode codepoint.
func vbsCompatChrW(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	return NewString(string(rune(args[0].Num))), nil
}

// vbsCompatMidB returns a byte-oriented substring with 1-based indexes.
func vbsCompatMidB(_ *VM, args []Value) (Value, error) {
	if len(args) < 2 {
		return NewString(""), nil
	}
	bs := ansiBytes(args[0].String())
	start := int(args[1].Num)
	if start < 1 {
		start = 1
	}
	start--
	if start >= len(bs) {
		return NewString(""), nil
	}
	end := len(bs)
	if len(args) > 2 {
		n := int(args[2].Num)
		if n < 0 {
			n = 0
		}
		end = start + n
		if end > len(bs) {
			end = len(bs)
		}
	}
	return NewString(bytesToVBByteString(bs[start:end])), nil
}

// vbsCompatLeftB returns left N bytes of ANSI representation.
func vbsCompatLeftB(_ *VM, args []Value) (Value, error) {
	if len(args) < 2 {
		return NewString(""), nil
	}
	bs := ansiBytes(args[0].String())
	n := int(args[1].Num)
	if n < 0 {
		n = 0
	}
	if n > len(bs) {
		n = len(bs)
	}
	return NewString(bytesToVBByteString(bs[:n])), nil
}

// vbsCompatRightB returns right N bytes of ANSI representation.
func vbsCompatRightB(_ *VM, args []Value) (Value, error) {
	if len(args) < 2 {
		return NewString(""), nil
	}
	bs := ansiBytes(args[0].String())
	n := int(args[1].Num)
	if n < 0 {
		n = 0
	}
	if n > len(bs) {
		n = len(bs)
	}
	return NewString(bytesToVBByteString(bs[len(bs)-n:])), nil
}

// vbsCompatInStrB searches byte position using 1-based semantics.
func vbsCompatInStrB(_ *VM, args []Value) (Value, error) {
	if len(args) < 2 {
		return NewInteger(0), nil
	}
	start := 1
	s1Idx := 0
	s2Idx := 1
	if len(args) >= 3 {
		start = int(args[0].Num)
		s1Idx = 1
		s2Idx = 2
	}
	bs1 := ansiBytes(args[s1Idx].String())
	bs2 := ansiBytes(args[s2Idx].String())
	if len(bs2) == 0 {
		return NewInteger(int64(start)), nil
	}
	startPos := start - 1
	if startPos < 0 {
		startPos = 0
	}
	if startPos >= len(bs1) {
		return NewInteger(0), nil
	}
	idx := bytes.Index(bs1[startPos:], bs2)
	if idx < 0 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(startPos + idx + 1)), nil
}

// vbsCompatHex converts a number to uppercase hexadecimal.
// VBScript rounds the argument to Long integer (banker's rounding) before converting.
func vbsCompatHex(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	var intVal int64
	switch args[0].Type {
	case VTDouble:
		intVal = int64(math.RoundToEven(args[0].Flt))
	case VTString:
		if parsed, _, ok := vbsParseNumericString(args[0].Str); ok {
			intVal = int64(math.RoundToEven(parsed))
		}
	default:
		intVal = args[0].Num
	}
	return NewString(fmt.Sprintf("%X", uint32(intVal))), nil
}

// vbsCompatInStrRev searches from right to left using 1-based semantics.
func vbsCompatInStrRev(_ *VM, args []Value) (Value, error) {
	if len(args) < 2 {
		return NewInteger(0), nil
	}
	s1 := args[0].String()
	s2 := args[1].String()
	runes1 := []rune(s1)
	start := len(runes1)
	compare := 0
	if len(args) >= 3 {
		v := int(args[2].Num)
		if v > 0 {
			start = v
		}
	}
	if len(args) >= 4 {
		compare = int(args[3].Num)
	}
	if start > len(runes1) {
		start = len(runes1)
	}
	if start < 1 || s2 == "" {
		return NewInteger(0), nil
	}
	search := string(runes1[:start])
	target := s2
	if compare == 1 {
		search = strings.ToLower(search)
		target = strings.ToLower(target)
	}
	idx := strings.LastIndex(search, target)
	if idx < 0 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(len([]rune(search[:idx])) + 1)), nil
}

// vbsCompatSplit splits a string into a zero-based Variant array.
func vbsCompatSplit(_ *VM, args []Value) (Value, error) {
	if len(args) < 2 {
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}, nil
	}
	s := args[0].String()
	delim := args[1].String()
	if delim == "" {
		runes := []rune(s)
		values := make([]Value, len(runes))
		for i := 0; i < len(runes); i++ {
			values[i] = NewString(string(runes[i]))
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}, nil
	}
	limit := -1
	if len(args) > 2 {
		limit = int(args[2].Num)
	}
	parts := []string{}
	if limit > 0 {
		parts = strings.SplitN(s, delim, limit)
	} else {
		parts = strings.Split(s, delim)
	}
	values := make([]Value, len(parts))
	for i := 0; i < len(parts); i++ {
		values[i] = NewString(parts[i])
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}, nil
}

// vbsCompatJoin joins one Variant array using the specified delimiter.
func vbsCompatJoin(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 || args[0].Type != VTArray || args[0].Arr == nil {
		return NewString(""), nil
	}
	delim := " "
	if len(args) > 1 {
		delim = args[1].String()
	}
	parts := make([]string, len(args[0].Arr.Values))
	for i := 0; i < len(parts); i++ {
		parts[i] = args[0].Arr.Values[i].String()
	}
	return NewString(strings.Join(parts, delim)), nil
}

// vbsCompatFilter filters string array elements by substring match.
func vbsCompatFilter(_ *VM, args []Value) (Value, error) {
	if len(args) < 2 || args[0].Type != VTArray || args[0].Arr == nil {
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}, nil
	}
	needle := args[1].String()
	include := true
	if len(args) > 2 {
		include = args[2].Num != 0
	}
	compare := 0
	if len(args) > 3 {
		compare = int(args[3].Num)
	}
	if compare == 1 {
		needle = strings.ToLower(needle)
	}
	matches := make([]Value, 0, len(args[0].Arr.Values))
	for i := 0; i < len(args[0].Arr.Values); i++ {
		candidate := args[0].Arr.Values[i].String()
		sample := candidate
		if compare == 1 {
			sample = strings.ToLower(sample)
		}
		has := strings.Contains(sample, needle)
		if (include && has) || (!include && !has) {
			matches = append(matches, NewString(candidate))
		}
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, matches)}, nil
}

// vbsCompatOct converts a number to octal.
// VBScript rounds the argument to Long integer (banker's rounding) before converting.
func vbsCompatOct(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	var intVal int64
	switch args[0].Type {
	case VTDouble:
		intVal = int64(math.RoundToEven(args[0].Flt))
	case VTString:
		if parsed, _, ok := vbsParseNumericString(args[0].Str); ok {
			intVal = int64(math.RoundToEven(parsed))
		}
	default:
		intVal = args[0].Num
	}
	return NewString(strconv.FormatUint(uint64(uint32(intVal)), 8)), nil
}

// vbsCompatRound rounds a number with optional decimal digits.
func vbsCompatRound(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewDouble(0), nil
	}
	value := float64(args[0].Num)
	if args[0].Type == VTDouble {
		value = args[0].Flt
	}
	digits := 0
	if len(args) > 1 {
		digits = int(args[1].Num)
	}
	factor := math.Pow(10, float64(digits))
	if factor == 0 || math.IsInf(factor, 0) || math.IsNaN(factor) {
		return NewDouble(0), nil
	}
	return NewDouble(math.RoundToEven(value*factor) / factor), nil
}

// vbsCompatSgn returns -1, 0, or 1 based on numeric sign.
func vbsCompatSgn(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	value := float64(args[0].Num)
	if args[0].Type == VTDouble {
		value = args[0].Flt
	}
	if value > 0 {
		return NewInteger(1), nil
	}
	if value < 0 {
		return NewInteger(-1), nil
	}
	return NewInteger(0), nil
}

// vbsCompatNow returns current date and time in configured location.
func vbsCompatNow(vm *VM, _ []Value) (Value, error) {
	return NewDate(time.Now().In(builtinCurrentLocation(vm))), nil
}

// vbsCompatDate returns current date with zeroed clock.
func vbsCompatDate(vm *VM, _ []Value) (Value, error) {
	now := time.Now().In(builtinCurrentLocation(vm))
	return NewDate(time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, now.Location())), nil
}

// vbsCompatTime returns current time anchored to VBScript base date.
func vbsCompatTime(vm *VM, _ []Value) (Value, error) {
	now := time.Now().In(builtinCurrentLocation(vm))
	return NewDate(time.Date(1899, time.December, 30, now.Hour(), now.Minute(), now.Second(), 0, now.Location())), nil
}

// vbsCompatYear returns the year part from a date value.
func vbsCompatYear(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(valueToTimeInLocale(vm, resolveCallable(vm, args[0])).In(builtinCurrentLocation(vm)).Year())), nil
}

// vbsCompatMonth returns the month part from a date value.
func vbsCompatMonth(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(valueToTimeInLocale(vm, resolveCallable(vm, args[0])).In(builtinCurrentLocation(vm)).Month())), nil
}

// vbsCompatDay returns the day part from a date value.
func vbsCompatDay(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(valueToTimeInLocale(vm, resolveCallable(vm, args[0])).In(builtinCurrentLocation(vm)).Day())), nil
}

// vbsCompatHour returns the hour part from a date value.
func vbsCompatHour(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(valueToTimeInLocale(vm, resolveCallable(vm, args[0])).In(builtinCurrentLocation(vm)).Hour())), nil
}

// vbsCompatMinute returns the minute part from a date value.
func vbsCompatMinute(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(valueToTimeInLocale(vm, resolveCallable(vm, args[0])).In(builtinCurrentLocation(vm)).Minute())), nil
}

// vbsCompatSecond returns the second part from a date value.
func vbsCompatSecond(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(valueToTimeInLocale(vm, resolveCallable(vm, args[0])).In(builtinCurrentLocation(vm)).Second())), nil
}

// vbsCompatGetLocale returns effective LCID used by current request.
func vbsCompatGetLocale(vm *VM, _ []Value) (Value, error) {
	return NewInteger(int64(builtinCurrentLCID(vm))), nil
}

// vbsCompatSetLocale updates request/session LCID and returns previous value.
func vbsCompatSetLocale(vm *VM, args []Value) (Value, error) {
	previous := builtinCurrentLCID(vm)
	if len(args) < 1 {
		return NewInteger(int64(previous)), nil
	}
	if vm != nil && vm.host != nil && vm.host.SessionEnabled() {
		vm.host.Session().SetLCID(int(args[0].Num))
	}
	return NewInteger(int64(previous)), nil
}

// vbsCompatMonthName returns month name with optional abbreviation and locale fallback.
func vbsCompatMonthName(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	month := int(args[0].Num)
	abbrev := len(args) > 1 && args[1].Num != 0
	if month < 1 || month > 12 {
		return NewString(""), nil
	}
	names := localizedMonthNames(builtinLocaleTag(vm), abbrev)
	return NewString(names[month-1]), nil
}

// vbsCompatWeekdayName returns weekday name honoring VBScript numbering.
func vbsCompatWeekdayName(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	weekday := int(args[0].Num)
	abbrev := len(args) > 1 && args[1].Num != 0
	if weekday < 1 || weekday > 7 {
		return NewString(""), nil
	}
	names := localizedWeekdayNames(builtinLocaleTag(vm), abbrev)
	return NewString(names[weekday-1]), nil
}

// resolveCallable handles VBScript's zero-arg auto-call semantics.
// In VBScript, writing "Now" inside an expression (e.g. FormatDateTime(Now, 1))
// is identical to calling "Now()" — the runtime evaluates zero-argument functions
// implicitly. When such a name is used without parentheses the compiler emits an
// OpGetGlobal that loads a VTBuiltin value onto the stack. This helper detects that
// case and invokes the built-in with an empty argument list so callers always receive
// the function's return value rather than the raw function reference.
func resolveCallable(vm *VM, v Value) Value {
	if v.Type != VTBuiltin {
		return v
	}
	idx := int(v.Num)
	if idx < 0 || idx >= len(BuiltinRegistry) {
		return v
	}
	result, err := BuiltinRegistry[idx](vm, nil)
	if err != nil {
		return v
	}
	return result
}

// valueToTimeInLocale converts a value to time.Time using locale-aware parsing rules.
func valueToTimeInLocale(vm *VM, v Value) time.Time {
	if v.Type == VTDate {
		if v.Num == vbsCompatZeroDateSentinelNum() {
			return time.Time{}
		}
		return time.Unix(0, v.Num)
	}
	text := strings.TrimSpace(v.String())
	if text == "" {
		return time.Time{}
	}
	return parseLocalizedTimeValue(text, builtinCurrentLocation(vm), builtinLocaleProfileForVM(vm))
}

// vbsCompatFormatDateTime formats date/time using locale-sensitive defaults.
func vbsCompatFormatDateTime(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	profile := builtinLocaleProfileForVM(vm)
	dateValue := valueToTimeInLocale(vm, resolveCallable(vm, args[0])).In(builtinCurrentLocation(vm))
	formatType := 0
	if len(args) > 1 {
		formatType = int(args[1].Num)
	}
	return NewString(localizedDateTimeText(dateValue, profile, formatType)), nil
}

// vbsCompatFormatNumber formats numeric values with fixed decimal places.
func vbsCompatFormatNumber(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString("0"), nil
	}
	value := float64(args[0].Num)
	if args[0].Type == VTDouble {
		value = args[0].Flt
	}
	digits := 2
	if len(args) > 1 {
		digits = int(args[1].Num)
	}
	useGrouping := true
	if len(args) > 2 && args[2].Type == VTInteger && args[2].Num == 0 {
		useGrouping = false
	}
	return NewString(localizedNumberString(value, digits, builtinLocaleProfileForVM(vm), useGrouping)), nil
}

// vbsCompatFormatCurrency formats currency values with locale-specific symbol fallback.
func vbsCompatFormatCurrency(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(localizedCurrencyString(0, 2, builtinLocaleProfileForVM(vm))), nil
	}
	value := float64(args[0].Num)
	if args[0].Type == VTDouble {
		value = args[0].Flt
	}
	digits := 2
	if len(args) > 1 {
		digits = int(args[1].Num)
	}
	return NewString(localizedCurrencyString(value, digits, builtinLocaleProfileForVM(vm))), nil
}

// vbsCompatFormatPercent formats percentage values with optional decimal precision.
func vbsCompatFormatPercent(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString("0%"), nil
	}
	value := float64(args[0].Num)
	if args[0].Type == VTDouble {
		value = args[0].Flt
	}
	digits := 2
	if len(args) > 1 {
		digits = int(args[1].Num)
	}
	formatted := localizedNumberString(value*100, digits, builtinLocaleProfileForVM(vm), false)
	return NewString(formatted + "%"), nil
}

// vbsCompatTimer returns the number of seconds elapsed since midnight.
func vbsCompatTimer(vm *VM, _ []Value) (Value, error) {
	now := time.Now().In(builtinCurrentLocation(vm))
	midnight := time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, now.Location())
	return NewDouble(now.Sub(midnight).Seconds()), nil
}

// vbsCompatIsNumeric tests whether a value can be interpreted as numeric.
func vbsCompatIsNumeric(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewBool(false), nil
	}
	switch args[0].Type {
	case VTInteger, VTDouble, VTBool:
		return NewBool(true), nil
	case VTString:
		text := strings.TrimSpace(args[0].Str)
		if text == "" {
			return NewBool(false), nil
		}
		if _, err := strconv.ParseInt(text, 10, 64); err == nil {
			return NewBool(true), nil
		}
		if _, err := strconv.ParseFloat(text, 64); err == nil {
			return NewBool(true), nil
		}
	}
	return NewBool(false), nil
}

// vbsCompatIsDate tests whether a value can be interpreted as date/time.
func vbsCompatIsDate(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewBool(false), nil
	}
	v := resolveCallable(vm, args[0])
	if v.Type == VTDate {
		return NewBool(!valueToTimeInLocale(vm, v).IsZero()), nil
	}
	return NewBool(!valueToTimeInLocale(vm, v).IsZero()), nil
}

// vbsCompatCreateObject exposes VBScript CreateObject built-in for compatibility.
func vbsCompatCreateObject(vm *VM, args []Value) (Value, error) {
	if vm == nil || vm.host == nil || len(args) < 1 {
		return NewEmpty(), nil
	}
	return vm.dispatchNativeCall(nativeObjectServer, "CreateObject", args), nil
}

// vbsCompatGetObject maps GetObject to CreateObject compatibility behavior.
func vbsCompatGetObject(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewEmpty(), nil
	}
	if len(args) >= 2 {
		return vbsCompatCreateObject(vm, []Value{args[1]})
	}
	return vbsCompatCreateObject(vm, args)
}

// vbsCompatUnescape decodes %XX and %uXXXX escapes.
func vbsCompatUnescape(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	raw := args[0].String()
	decoded, err := url.QueryUnescape(strings.ReplaceAll(raw, "+", "%2B"))
	if err == nil {
		raw = decoded
	}
	raw = strings.ReplaceAll(raw, "%u", "\\u")
	runes := []rune{}
	for i := 0; i < len(raw); i++ {
		if i+5 < len(raw) && raw[i] == '\\' && raw[i+1] == 'u' {
			hexCode := raw[i+2 : i+6]
			if n, parseErr := strconv.ParseUint(hexCode, 16, 16); parseErr == nil {
				runes = append(runes, rune(n))
				i += 5
				continue
			}
		}
		runes = append(runes, rune(raw[i]))
	}
	return NewString(string(runes)), nil
}

// vbsCompatDateSerial builds a date value from year, month, and day.
func vbsCompatDateSerial(vm *VM, args []Value) (Value, error) {
	if len(args) < 3 {
		return NewDate(time.Time{}), nil
	}
	y, ok := vbsCompatCoerceSerialInt(vm, args[0])
	if !ok {
		vm.raise(vbscript.TypeMismatch, "Type mismatch: DateSerial year must be numeric")
		return NewEmpty(), nil
	}
	mInt, ok := vbsCompatCoerceSerialInt(vm, args[1])
	if !ok {
		vm.raise(vbscript.TypeMismatch, "Type mismatch: DateSerial month must be numeric")
		return NewEmpty(), nil
	}
	d, ok := vbsCompatCoerceSerialInt(vm, args[2])
	if !ok {
		vm.raise(vbscript.TypeMismatch, "Type mismatch: DateSerial day must be numeric")
		return NewEmpty(), nil
	}
	loc := builtinCurrentLocation(vm)
	m := time.Month(mInt)
	return NewDate(time.Date(y, m, d, 0, 0, 0, 0, loc)), nil
}

// vbsCompatTimeSerial builds a time value from hour, minute, and second.
func vbsCompatTimeSerial(vm *VM, args []Value) (Value, error) {
	if len(args) < 3 {
		return NewDate(time.Time{}), nil
	}
	h, ok := vbsCompatCoerceSerialInt(vm, args[0])
	if !ok {
		vm.raise(vbscript.TypeMismatch, "Type mismatch: TimeSerial hour must be numeric")
		return NewEmpty(), nil
	}
	m, ok := vbsCompatCoerceSerialInt(vm, args[1])
	if !ok {
		vm.raise(vbscript.TypeMismatch, "Type mismatch: TimeSerial minute must be numeric")
		return NewEmpty(), nil
	}
	s, ok := vbsCompatCoerceSerialInt(vm, args[2])
	if !ok {
		vm.raise(vbscript.TypeMismatch, "Type mismatch: TimeSerial second must be numeric")
		return NewEmpty(), nil
	}
	loc := builtinCurrentLocation(vm)
	return NewDate(time.Date(1899, time.December, 30, h, m, s, 0, loc)), nil
}

// vbsCompatDateValue parses a date string.
func vbsCompatDateValue(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewDate(time.Time{}), nil
	}
	parsed := valueToTimeInLocale(vm, NewString(args[0].String())).In(builtinCurrentLocation(vm))
	return NewDate(time.Date(parsed.Year(), parsed.Month(), parsed.Day(), 0, 0, 0, 0, parsed.Location())), nil
}

// vbsCompatTimeValue parses a time string.
func vbsCompatTimeValue(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewDate(time.Time{}), nil
	}
	parsed := valueToTimeInLocale(vm, NewString(args[0].String())).In(builtinCurrentLocation(vm))
	return NewDate(time.Date(1899, time.December, 30, parsed.Hour(), parsed.Minute(), parsed.Second(), 0, parsed.Location())), nil
}

// vbsCompatWeekday returns weekday number using Sunday=1 semantics.
func vbsCompatWeekday(vm *VM, args []Value) (Value, error) {
	t := time.Now()
	if len(args) > 0 {
		t = valueToTimeInLocale(vm, resolveCallable(vm, args[0])).In(builtinCurrentLocation(vm))
	} else {
		t = t.In(builtinCurrentLocation(vm))
	}
	return NewInteger(int64(int(t.Weekday()) + 1)), nil
}

// vbsCompatDateAdd adds an interval to a date value.
func vbsCompatDateAdd(vm *VM, args []Value) (Value, error) {
	if len(args) < 3 {
		return NewDate(time.Time{}), nil
	}
	interval := strings.ToLower(args[0].String())
	number := int(args[1].Num)
	value := valueToTimeInLocale(vm, resolveCallable(vm, args[2]))
	switch interval {
	case "yyyy":
		value = value.AddDate(number, 0, 0)
	case "m":
		value = value.AddDate(0, number, 0)
	case "d", "y", "w":
		value = value.AddDate(0, 0, number)
	case "h":
		value = value.Add(time.Duration(number) * time.Hour)
	case "n":
		value = value.Add(time.Duration(number) * time.Minute)
	case "s":
		value = value.Add(time.Duration(number) * time.Second)
	}
	return NewDate(value), nil
}

// vbsCompatDateDiff returns interval difference between two date values.
func vbsCompatDateDiff(vm *VM, args []Value) (Value, error) {
	if len(args) < 3 {
		return NewInteger(0), nil
	}
	interval := strings.ToLower(args[0].String())
	start := valueToTimeInLocale(vm, resolveCallable(vm, args[1])).In(builtinCurrentLocation(vm))
	end := valueToTimeInLocale(vm, resolveCallable(vm, args[2])).In(builtinCurrentLocation(vm))
	delta := end.Sub(start)
	startCalendarDay := time.Date(start.Year(), start.Month(), start.Day(), 0, 0, 0, 0, time.UTC)
	endCalendarDay := time.Date(end.Year(), end.Month(), end.Day(), 0, 0, 0, 0, time.UTC)
	dayDiff := int64(endCalendarDay.Sub(startCalendarDay) / (24 * time.Hour))
	switch interval {
	case "yyyy":
		return NewInteger(int64(end.Year() - start.Year())), nil
	case "m":
		return NewInteger(int64((end.Year()-start.Year())*12 + int(end.Month()-start.Month()))), nil
	case "d", "y", "w":
		return NewInteger(dayDiff), nil
	case "h":
		return NewInteger(int64(delta.Hours())), nil
	case "n":
		return NewInteger(int64(delta.Minutes())), nil
	case "s":
		return NewInteger(int64(delta.Seconds())), nil
	default:
		return NewInteger(0), nil
	}
}

// vbsCompatDatePart returns one specific part from a date value.
func vbsCompatDatePart(vm *VM, args []Value) (Value, error) {
	if len(args) < 2 {
		return NewInteger(0), nil
	}
	interval := strings.ToLower(args[0].String())
	value := valueToTimeInLocale(vm, resolveCallable(vm, args[1])).In(builtinCurrentLocation(vm))
	switch interval {
	case "yyyy":
		return NewInteger(int64(value.Year())), nil
	case "m":
		return NewInteger(int64(value.Month())), nil
	case "d", "y":
		return NewInteger(int64(value.Day())), nil
	case "h":
		return NewInteger(int64(value.Hour())), nil
	case "n":
		return NewInteger(int64(value.Minute())), nil
	case "s":
		return NewInteger(int64(value.Second())), nil
	case "w":
		return NewInteger(int64(int(value.Weekday()) + 1)), nil
	default:
		return NewInteger(0), nil
	}
}

// vbsCompatCBool converts a value to VBScript boolean semantics.
func vbsCompatCBool(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewBool(false), nil
	}
	v := args[0]
	switch v.Type {
	case VTBool:
		return v, nil
	case VTInteger:
		return NewBool(v.Num != 0), nil
	case VTDouble:
		return NewBool(v.Flt != 0), nil
	case VTString:
		t := strings.TrimSpace(strings.ToLower(v.Str))
		return NewBool(!(t == "" || t == "0" || t == "false")), nil
	default:
		return NewBool(false), nil
	}
}

// vbsCompatCLng converts a value to 32-bit long with VBScript rounding and errors.
func vbsCompatCLng(vm *VM, args []Value) (Value, error) {
	return vbsConvertRoundedIntegerVM(vm, args, -2147483648, 2147483647)
}

// vbsCompatCSng converts a value to single-precision numeric result.
func vbsCompatCSng(vm *VM, args []Value) (Value, error) {
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

	if numericValue > math.MaxFloat32 || numericValue < -math.MaxFloat32 {
		if vm != nil {
			vm.raise(vbscript.Overflow, vbscript.Overflow.String())
		}
		return NewDouble(0), nil
	}

	return NewDouble(float64(float32(numericValue))), nil
}

// vbsCompatCByte converts a value to byte range 0..255.
func vbsCompatCByte(vm *VM, args []Value) (Value, error) {
	return vbsConvertRoundedIntegerVM(vm, args, 0, 255)
}

// vbsCompatCDate converts a value to date.
func vbsCompatCDate(vm *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewDate(time.Time{}), nil
	}
	return NewDate(valueToTimeInLocale(vm, args[0])), nil
}

// vbsCompatCChar converts to one-character string.
func vbsCompatCChar(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewString(""), nil
	}
	if args[0].Type == VTString {
		r := []rune(args[0].Str)
		if len(r) == 0 {
			return NewString(""), nil
		}
		return NewString(string(r[0])), nil
	}
	return NewString(string(rune(args[0].Num))), nil
}

// vbsCompatCCur converts to currency-compatible floating-point value.
func vbsCompatCCur(vm *VM, args []Value) (Value, error) {
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

// vbsCompatCDec converts to decimal-compatible floating-point value.
func vbsCompatCDec(vm *VM, args []Value) (Value, error) {
	return vbsCompatCCur(vm, args)
}

// vbsCompatCObj keeps object conversion semantics by returning the original value.
func vbsCompatCObj(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewEmpty(), nil
	}
	return args[0], nil
}

// vbsCompatCVar keeps variant conversion semantics by returning the original value.
func vbsCompatCVar(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewEmpty(), nil
	}
	return args[0], nil
}

// vbsCompatCSByte converts to signed byte range.
func vbsCompatCSByte(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	value := args[0].Num & 0xFF
	if value > 127 {
		value = value - 256
	}
	return NewInteger(value), nil
}

// vbsCompatCShort converts to signed 16-bit integer range.
func vbsCompatCShort(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	value := int64(int16(args[0].Num))
	return NewInteger(value), nil
}

// vbsCompatCUInt converts to unsigned 32-bit integer stored in VTInteger.
func vbsCompatCUInt(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(uint32(args[0].Num))), nil
}

// vbsCompatCULng converts to unsigned 64-bit integer clipped to int64 space.
func vbsCompatCULng(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	if args[0].Num < 0 {
		return NewInteger(0), nil
	}
	return NewInteger(args[0].Num), nil
}

// vbsCompatCUShort converts to unsigned 16-bit integer stored in VTInteger.
func vbsCompatCUShort(_ *VM, args []Value) (Value, error) {
	if len(args) < 1 {
		return NewInteger(0), nil
	}
	return NewInteger(int64(uint16(args[0].Num))), nil
}

// vbsCompatSin computes sine in radians.
func vbsCompatSin(_ *VM, args []Value) (Value, error) { return trigUnary(args, math.Sin) }

// vbsCompatCos computes cosine in radians.
func vbsCompatCos(_ *VM, args []Value) (Value, error) { return trigUnary(args, math.Cos) }

// vbsCompatTan computes tangent in radians.
func vbsCompatTan(_ *VM, args []Value) (Value, error) { return trigUnary(args, math.Tan) }

// vbsCompatAtn computes arctangent in radians.
func vbsCompatAtn(_ *VM, args []Value) (Value, error) { return trigUnary(args, math.Atan) }

// vbsCompatLog computes natural logarithm.
func vbsCompatLog(_ *VM, args []Value) (Value, error) { return trigUnary(args, math.Log) }

// vbsCompatExp computes e raised to the argument.
func vbsCompatExp(_ *VM, args []Value) (Value, error) { return trigUnary(args, math.Exp) }

// trigUnary applies one float64 unary math function to a VM argument.
func trigUnary(args []Value, fn func(float64) float64) (Value, error) {
	if len(args) < 1 {
		return NewDouble(0), nil
	}
	input := float64(args[0].Num)
	if args[0].Type == VTDouble {
		input = args[0].Flt
	}
	return NewDouble(fn(input)), nil
}

// vbsCompatExecuteGlobal compiles and executes code in the active global scope.
func vbsCompatExecuteGlobal(vm *VM, args []Value) (Value, error) {
	if vm == nil || len(args) < 1 {
		return NewEmpty(), nil
	}
	callerOnResumeNext := vm.onResumeNext

	code := vm.valueToString(resolveCallable(vm, args[0]))
	if strings.TrimSpace(code) == "" {
		return NewEmpty(), nil
	}

	compiled, err := vm.getOrCompileDynamicProgram(code, Value{}, dynamicExecKindExecuteGlobal)
	if err != nil {
		var syntaxErr *vbscript.VBSyntaxError
		if errors.As(err, &syntaxErr) {
			vm.raiseFromSyntaxError(syntaxErr)
			return NewEmpty(), nil
		}
		return NewEmpty(), err
	}
	if compiled == nil {
		return NewEmpty(), nil
	}

	startIP := vm.appendCachedDynamicProgram(compiled)
	vm.applyCompilerSnapshot(compiled)
	child := vm.cloneForExecuteGlobal(startIP)
	err = child.Run()
	vm.syncExecuteGlobalState(child)
	if err != nil {
		if callerOnResumeNext {
			if vme, ok := err.(*VMError); ok {
				vm.errSetFromVMError(vme)
				vm.lastError = vme
			} else {
				// Preserve Resume Next semantics: surface error via Err without aborting caller.
				vme := &VMError{
					Code:           vbscript.InternalError,
					Line:           vm.lastLine,
					Column:         vm.lastColumn,
					Msg:            err.Error(),
					ASPCode:        int(vbscript.InternalError),
					ASPDescription: err.Error(),
					Category:       "VBScript runtime",
					Description:    err.Error(),
					Number:         vbscript.HRESULTFromVBScriptCode(vbscript.InternalError),
					Source:         "VBScript runtime error",
				}
				vm.errSetFromVMError(vme)
				vm.lastError = vme
			}
			return NewEmpty(), nil
		}
		return NewEmpty(), err
	}

	return NewEmpty(), nil
}

// vbsCompatExecute compiles and executes dynamic code in the active runtime scope.
func vbsCompatExecute(vm *VM, args []Value) (Value, error) {
	if vm == nil || len(args) < 1 {
		return NewEmpty(), nil
	}

	code := vm.valueToString(resolveCallable(vm, args[0]))
	if strings.TrimSpace(code) == "" {
		return NewEmpty(), nil
	}

	var localSub Value
	if len(vm.callStack) > 0 {
		localSub = vm.callStack[len(vm.callStack)-1].callee
	}

	compiled, err := vm.getOrCompileDynamicProgram(code, localSub, dynamicExecKindExecute)
	if err != nil {
		var syntaxErr *vbscript.VBSyntaxError
		if errors.As(err, &syntaxErr) {
			vm.raiseFromSyntaxError(syntaxErr)
			return NewEmpty(), nil
		}
		return NewEmpty(), err
	}
	if compiled == nil {
		return NewEmpty(), nil
	}

	startIP := vm.appendCachedDynamicProgram(compiled)
	vm.applyCompilerSnapshot(compiled)

	// Clone VM sharing parent stack and FP/SP context
	child := vm.cloneForExecuteLocal(startIP)

	err = child.Run()
	// Sync stack and SP back after execution
	vm.syncExecuteLocalState(child)

	if err != nil {
		return NewEmpty(), err
	}

	return NewEmpty(), nil
}

// vbsCompatGetRef resolves one function/sub reference by name.
func vbsCompatGetRef(vm *VM, args []Value) (Value, error) {
	if vm == nil || len(args) < 1 {
		return NewEmpty(), nil
	}

	name := strings.TrimSpace(vm.valueToString(resolveCallable(vm, args[0])))
	if strings.HasSuffix(name, "()") {
		name = strings.TrimSpace(strings.TrimSuffix(name, "()"))
	}
	if name == "" {
		return NewEmpty(), nil
	}
	lowerName := strings.ToLower(name)

	if idx, ok := vm.globalNameIndex[lowerName]; ok {
		if idx >= 0 && idx < len(vm.Globals) {
			resolved := vm.Globals[idx]
			if resolved.Type == VTUserSub || resolved.Type == VTBuiltin {
				return resolved, nil
			}
		}
	}

	if builtinIdx, ok := GetBuiltinIndex(name); ok {
		return Value{Type: VTBuiltin, Num: int64(builtinIdx)}, nil
	}

	return NewEmpty(), nil
}

type evalParser struct {
	text string
	pos  int
}

func (p *evalParser) skipSpaces() {
	for p.pos < len(p.text) {
		ch := p.text[p.pos]
		if ch != ' ' && ch != '\t' && ch != '\r' && ch != '\n' {
			break
		}
		p.pos++
	}
}

func (p *evalParser) parseNumber() (float64, bool, bool) {
	p.skipSpaces()
	start := p.pos
	hasDot := false
	hasExp := false

	for p.pos < len(p.text) {
		ch := p.text[p.pos]
		if ch >= '0' && ch <= '9' {
			p.pos++
			continue
		}
		if ch == '.' && !hasDot && !hasExp {
			hasDot = true
			p.pos++
			continue
		}
		if (ch == 'e' || ch == 'E') && !hasExp {
			hasExp = true
			p.pos++
			if p.pos < len(p.text) && (p.text[p.pos] == '+' || p.text[p.pos] == '-') {
				p.pos++
			}
			continue
		}
		break
	}

	if p.pos == start {
		return 0, false, false
	}

	n, err := strconv.ParseFloat(p.text[start:p.pos], 64)
	if err != nil {
		return 0, false, false
	}

	return n, hasDot || hasExp, true
}

func (p *evalParser) parseFactor() (float64, bool, bool) {
	p.skipSpaces()
	if p.pos >= len(p.text) {
		return 0, false, false
	}

	if p.text[p.pos] == '+' {
		p.pos++
		return p.parseFactor()
	}
	if p.text[p.pos] == '-' {
		p.pos++
		v, isFloat, ok := p.parseFactor()
		if !ok {
			return 0, false, false
		}
		return -v, isFloat, true
	}

	if p.text[p.pos] == '(' {
		p.pos++
		v, isFloat, ok := p.parseExpr()
		if !ok {
			return 0, false, false
		}
		p.skipSpaces()
		if p.pos >= len(p.text) || p.text[p.pos] != ')' {
			return 0, false, false
		}
		p.pos++
		return v, isFloat, true
	}

	return p.parseNumber()
}

func (p *evalParser) parseTerm() (float64, bool, bool) {
	left, isFloat, ok := p.parseFactor()
	if !ok {
		return 0, false, false
	}

	for {
		p.skipSpaces()
		if p.pos >= len(p.text) {
			break
		}

		op := p.text[p.pos]
		if op != '*' && op != '/' {
			break
		}
		p.pos++

		right, rightFloat, rok := p.parseFactor()
		if !rok {
			return 0, false, false
		}

		if op == '*' {
			left *= right
			isFloat = isFloat || rightFloat
		} else {
			if right == 0 {
				return 0, false, false
			}
			left /= right
			isFloat = true
		}
	}

	return left, isFloat, true
}

func (p *evalParser) parseExpr() (float64, bool, bool) {
	left, isFloat, ok := p.parseTerm()
	if !ok {
		return 0, false, false
	}

	for {
		p.skipSpaces()
		if p.pos >= len(p.text) {
			break
		}

		op := p.text[p.pos]
		if op != '+' && op != '-' {
			break
		}
		p.pos++

		right, rightFloat, rok := p.parseTerm()
		if !rok {
			return 0, false, false
		}

		if op == '+' {
			left += right
		} else {
			left -= right
		}
		isFloat = isFloat || rightFloat
	}

	return left, isFloat, true
}

// vbsCompatEval evaluates a VBScript expression string with full access to current VM scope.
// It supports variables, functions, arrays, and all operators - achieving full compatibility
// with Classic ASP Eval semantics. Implementation compiles the expression directly
// and executes it in the current scope, capturing the stack result.
func vbsCompatEval(vm *VM, args []Value) (Value, error) {
	if vm == nil || len(args) < 1 {
		return NewEmpty(), nil
	}

	expr := strings.TrimSpace(vm.valueToString(resolveCallable(vm, args[0])))
	expr = strings.TrimLeft(expr, "\uFEFF")
	if expr == "" {
		return NewEmpty(), nil
	}

	var localSub Value
	if len(vm.callStack) > 0 {
		localSub = vm.callStack[len(vm.callStack)-1].callee
	}

	compiled, err := vm.getOrCompileEvalProgram(expr, localSub)
	if err != nil {
		var syntaxErr *vbscript.VBSyntaxError
		if errors.As(err, &syntaxErr) {
			vm.raiseFromSyntaxError(syntaxErr)
			return NewEmpty(), nil
		}
		return NewEmpty(), err
	}

	if compiled == nil {
		return NewEmpty(), nil
	}

	// Append one cloned compiled program to VM bytecode.
	startIP := vm.appendExecuteProgram(compiled.globalCount, compiled.constants, compiled.bytecode)

	if startIP < 0 || startIP >= len(vm.bytecode) {
		return NewEmpty(), nil
	}

	// Clone VM sharing parent stack and FP/SP context
	child := vm.cloneForExecuteLocal(startIP)

	// Execute the compiled expression
	if err := child.Run(); err != nil {
		// Eval must not overwrite the caller stack/frame state. We only sync
		// mutable runtime/global state from the child execution.
		vm.syncExecuteGlobalState(child)
		return NewEmpty(), err
	}

	// Capture the expression result from the top of the child stack before syncing.
	var resultValue Value
	if child.sp >= 0 {
		resultValue = child.stack[child.sp]
	} else {
		resultValue = NewEmpty()
	}

	// Sync only mutable runtime/global state. Preserving caller stack/frame keeps
	// Eval side-effect free with respect to VM stack discipline.
	vm.syncExecuteGlobalState(child)

	return resultValue, nil
}

// vbsCompatStrConv supports VBScript-like case conversion constants.
func vbsCompatStrConv(_ *VM, args []Value) (Value, error) {
	if len(args) < 2 {
		return NewString(""), nil
	}
	input := args[0].String()
	mode := int(args[1].Num)
	switch mode {
	case 1:
		return NewString(strings.ToUpper(input)), nil
	case 2:
		return NewString(strings.ToLower(input)), nil
	case 3:
		runes := []rune(strings.ToLower(input))
		newWord := true
		for i := 0; i < len(runes); i++ {
			if unicode.IsSpace(runes[i]) {
				newWord = true
				continue
			}
			if newWord {
				runes[i] = unicode.ToUpper(runes[i])
				newWord = false
			}
		}
		return NewString(string(runes)), nil
	case 64:
		u16 := utf16.Encode([]rune(input))
		bytesOut := make([]byte, 0, len(u16)*2)
		for _, u := range u16 {
			bytesOut = append(bytesOut, byte(u), byte(u>>8))
		}
		return NewString(string(bytesOut)), nil
	default:
		return NewString(input), nil
	}
}

// init registers compatibility built-ins ported from the legacy VM implementation.
func init() {
	RegisterBuiltin("IsNothing", vbsCompatIsNothing)
	RegisterBuiltin("VarType", vbsCompatVarType)
	RegisterBuiltin("ScriptEngine", vbsCompatScriptEngine)
	RegisterBuiltin("ScriptEngineMajorVersion", vbsCompatScriptEngineMajorVersion)
	RegisterBuiltin("ScriptEngineMinorVersion", vbsCompatScriptEngineMinorVersion)
	RegisterBuiltin("ScriptEngineBuildVersion", vbsCompatScriptEngineBuildVersion)
	RegisterBuiltin("LTrim", vbsCompatLTrim)
	RegisterBuiltin("RTrim", vbsCompatRTrim)
	RegisterBuiltin("Space", vbsCompatSpace)
	RegisterBuiltin("String", vbsCompatString)
	RegisterBuiltin("StrReverse", vbsCompatStrReverse)
	RegisterBuiltin("StrComp", vbsCompatStrComp)
	RegisterBuiltin("InStrRev", vbsCompatInStrRev)
	RegisterBuiltin("Split", vbsCompatSplit)
	RegisterBuiltin("Join", vbsCompatJoin)
	RegisterBuiltin("Filter", vbsCompatFilter)
	RegisterBuiltin("AscB", vbsCompatAscB)
	RegisterBuiltin("ChrB", vbsCompatChrB)
	RegisterBuiltin("AscW", vbsCompatAscW)
	RegisterBuiltin("ChrW", vbsCompatChrW)
	RegisterBuiltin("MidB", vbsCompatMidB)
	RegisterBuiltin("LeftB", vbsCompatLeftB)
	RegisterBuiltin("RightB", vbsCompatRightB)
	RegisterBuiltin("InStrB", vbsCompatInStrB)
	RegisterBuiltin("Hex", vbsCompatHex)
	RegisterBuiltin("Oct", vbsCompatOct)
	RegisterBuiltin("Round", vbsCompatRound)
	RegisterBuiltin("Sgn", vbsCompatSgn)
	RegisterBuiltin("Sin", vbsCompatSin)
	RegisterBuiltin("Cos", vbsCompatCos)
	RegisterBuiltin("Tan", vbsCompatTan)
	RegisterBuiltin("Atn", vbsCompatAtn)
	RegisterBuiltin("Log", vbsCompatLog)
	RegisterBuiltin("Exp", vbsCompatExp)
	RegisterBuiltin("Now", vbsCompatNow)
	RegisterBuiltin("Date", vbsCompatDate)
	RegisterBuiltin("Time", vbsCompatTime)
	RegisterBuiltin("Year", vbsCompatYear)
	RegisterBuiltin("Month", vbsCompatMonth)
	RegisterBuiltin("Day", vbsCompatDay)
	RegisterBuiltin("Hour", vbsCompatHour)
	RegisterBuiltin("Minute", vbsCompatMinute)
	RegisterBuiltin("Second", vbsCompatSecond)
	RegisterBuiltin("GetLocale", vbsCompatGetLocale)
	RegisterBuiltin("SetLocale", vbsCompatSetLocale)
	RegisterBuiltin("MonthName", vbsCompatMonthName)
	RegisterBuiltin("WeekdayName", vbsCompatWeekdayName)
	RegisterBuiltin("FormatDateTime", vbsCompatFormatDateTime)
	RegisterBuiltin("FormatNumber", vbsCompatFormatNumber)
	RegisterBuiltin("FormatCurrency", vbsCompatFormatCurrency)
	RegisterBuiltin("FormatPercent", vbsCompatFormatPercent)
	RegisterBuiltin("Timer", vbsCompatTimer)
	RegisterBuiltin("IsNumeric", vbsCompatIsNumeric)
	RegisterBuiltin("IsDate", vbsCompatIsDate)
	RegisterBuiltin("CreateObject", vbsCompatCreateObject)
	RegisterBuiltin("GetObject", vbsCompatGetObject)
	RegisterBuiltin("Unescape", vbsCompatUnescape)
	RegisterBuiltin("DateSerial", vbsCompatDateSerial)
	RegisterBuiltin("TimeSerial", vbsCompatTimeSerial)
	RegisterBuiltin("DateValue", vbsCompatDateValue)
	RegisterBuiltin("TimeValue", vbsCompatTimeValue)
	RegisterBuiltin("Weekday", vbsCompatWeekday)
	RegisterBuiltin("DateAdd", vbsCompatDateAdd)
	RegisterBuiltin("DateDiff", vbsCompatDateDiff)
	RegisterBuiltin("DatePart", vbsCompatDatePart)
	RegisterBuiltin("CBool", vbsCompatCBool)
	RegisterBuiltin("CLng", vbsCompatCLng)
	RegisterBuiltin("CSng", vbsCompatCSng)
	RegisterBuiltin("CByte", vbsCompatCByte)
	RegisterBuiltin("CDate", vbsCompatCDate)
	RegisterBuiltin("CChar", vbsCompatCChar)
	RegisterBuiltin("CCur", vbsCompatCCur)
	RegisterBuiltin("CDec", vbsCompatCDec)
	RegisterBuiltin("CObj", vbsCompatCObj)
	RegisterBuiltin("CVar", vbsCompatCVar)
	RegisterBuiltin("CSByte", vbsCompatCSByte)
	RegisterBuiltin("CShort", vbsCompatCShort)
	RegisterBuiltin("CUInt", vbsCompatCUInt)
	RegisterBuiltin("CULng", vbsCompatCULng)
	RegisterBuiltin("CUShort", vbsCompatCUShort)
	RegisterBuiltin("StrConv", vbsCompatStrConv)
	RegisterBuiltin("Execute", vbsCompatExecute)
	RegisterBuiltin("ExecuteGlobal", vbsCompatExecuteGlobal)
	RegisterBuiltin("Eval", vbsCompatEval)
	RegisterBuiltin("GetRef", vbsCompatGetRef)
}
