/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda, Yuri Eisman (@yeisman)
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
	"math"
	"strconv"
	"strings"

	"g3pix.com.br/axonasp/v2/vbscript"
)

// isNull reports whether the value is VBScript Null.
func isNull(v Value) bool {
	return v.Type == VTNull
}

// isEmpty reports whether the value is VBScript Empty.
func isEmpty(v Value) bool {
	return v.Type == VTEmpty
}

// isString reports whether the underlying variant subtype is String.
func isString(v Value) bool {
	return v.Type == VTString
}

// isNumericLike reports whether the value participates in VBScript numeric addition.
func isNumericLike(v Value) bool {
	switch v.Type {
	case VTBool, VTInteger, VTCurrency, VTDouble, VTDate:
		return true
	default:
		return false
	}
}

// coerceFloatStrict converts a value to float64 using VBScript-style numeric coercion.
// String conversion failure returns false so callers can raise Type mismatch.
// A Date converts to its day serial (vbsDateSerial) here and in the coercions below.
func (vm *VM) coerceFloatStrict(v Value) (float64, bool) {
	v = vm.unwrapADODBFieldValue(resolveCallable(vm, v))
	switch v.Type {
	case VTEmpty:
		return 0, true
	case VTDouble:
		return v.Flt, true
	case VTCurrency:
		return float64(v.Num) / 10000.0, true
	case VTBool:
		return float64(vbsBoolNum(v)), true
	case VTDate:
		return vbsDateSerial(v), true
	case VTInteger, VTNativeObject, VTBuiltin:
		return float64(v.Num), true
	case VTString:
		parsed, _, ok := vbsParseNumericStringInLocale(v.Str, vm.localeProfile())
		return parsed, ok
	default:
		return 0, false
	}
}

// coerceFloat converts a VM value into float64 using VBScript-like numeric coercion.
func (vm *VM) coerceFloat(v Value) float64 {
	v = vm.unwrapADODBFieldValue(resolveCallable(vm, v))
	switch v.Type {
	case VTDouble:
		return v.Flt
	case VTCurrency:
		return float64(v.Num) / 10000.0
	case VTBool:
		return float64(vbsBoolNum(v))
	case VTDate:
		return vbsDateSerial(v)
	case VTString:
		parsed, _, _ := vbsParseNumericStringInLocale(v.Str, vm.localeProfile())
		return parsed
	default:
		return float64(v.Num)
	}
}

// coerceInt64 converts a VM value into int64 using VBScript-like numeric coercion.
// Float-to-integer conversions use Banker's Rounding (round-half-to-even) to match
// the behavior of VBScript's implicit type coercion for Integer and Long variables.
func (vm *VM) coerceInt64(v Value) int64 {
	v = vm.unwrapADODBFieldValue(resolveCallable(vm, v))
	switch v.Type {
	case VTBool:
		return vbsBoolNum(v)
	case VTCurrency:
		return int64(math.RoundToEven(float64(v.Num) / 10000.0))
	case VTDate:
		return int64(math.RoundToEven(vbsDateSerial(v)))
	case VTInteger, VTNativeObject, VTBuiltin:
		return v.Num
	case VTDouble:
		return int64(math.RoundToEven(v.Flt))
	case VTString:
		parsed, err := strconv.ParseInt(strings.TrimSpace(v.Str), 10, 64)
		if err == nil {
			return parsed
		}
		parsedFloat, _, _ := vbsParseNumericStringInLocale(v.Str, vm.localeProfile())
		return int64(math.RoundToEven(parsedFloat))
	default:
		return 0
	}
}

// coerceLogicalInt64 converts a VM value to int64 for use in logical/bitwise operators
// (Not, And, Or, Xor, Eqv, Imp). Unlike coerceInt64, it validates string operands:
// only numeric-looking strings are accepted. Non-numeric strings (including "True",
// "False", and empty string "") return (0, false), signalling that the caller must
// raise a VBScript Type mismatch error — matching original VBScript behaviour where
// these operators require CLng-compatible operands.
func (vm *VM) coerceLogicalInt64(v Value) (int64, bool) {
	v = vm.unwrapADODBFieldValue(resolveCallable(vm, v))
	switch v.Type {
	case VTBool:
		return vbsBoolNum(v), true
	case VTCurrency:
		return int64(math.RoundToEven(float64(v.Num) / 10000.0)), true
	case VTDate:
		return int64(math.RoundToEven(vbsDateSerial(v))), true
	case VTInteger, VTNativeObject, VTBuiltin:
		return v.Num, true
	case VTDouble:
		return int64(math.RoundToEven(v.Flt)), true
	case VTEmpty:
		return 0, true
	case VTString:
		text := strings.TrimSpace(v.Str)
		if parsed, err := strconv.ParseInt(text, 10, 64); err == nil {
			return parsed, true
		}
		if parsedFloat, _, ok := vbsParseNumericStringInLocale(text, vm.localeProfile()); ok {
			return int64(math.RoundToEven(parsedFloat)), true
		}
		return 0, false // non-numeric string — raises Type mismatch
	default:
		return 0, false
	}
}

// isVBScriptIntegerString reports whether a trimmed numeric string uses integer syntax
// (no decimal point or scientific exponent outside &H/&O literals).
func isVBScriptIntegerString(trimmed string) bool {
	return isVBScriptIntegerStringInLocale(trimmed, ".")
}

// newVBScriptIntValue returns VTInteger with Long subtype when outside 16-bit signed range or forced.
func newVBScriptIntValue(val int64, forceLong bool) Value {
	if forceLong || val < -32768 || val > 32767 {
		return NewLong(val)
	}
	return NewInteger(val)
}

// coerceNumericValue converts a VBScript value to a numeric subtype (VTInteger, VTCurrency,
// VTDouble, or VTDate) for numeric contexts such as For...Next loop bounds.
func (vm *VM) coerceNumericValue(v Value) (Value, vbscript.VBSyntaxErrorCode, bool) {
	v = vm.unwrapADODBFieldValue(resolveCallable(vm, v))
	switch v.Type {
	case VTInteger, VTCurrency, VTDouble, VTDate:
		return v, 0, true
	case VTEmpty:
		return NewInteger(0), 0, true
	case VTBool:
		return NewInteger(vbsBoolNum(v)), 0, true
	case VTNull:
		return NewEmpty(), vbscript.InvalidUseOfNull, false
	case VTString:
		trimmed := strings.TrimSpace(v.Str)
		if n, err := strconv.ParseInt(trimmed, 10, 64); err == nil {
			return newVBScriptIntValue(n, false), 0, true
		}
		profile := vm.localeProfile()
		parsed, code, ok := vbsParseNumericStringInLocale(trimmed, profile)
		if !ok {
			return NewEmpty(), code, false
		}
		if isVBScriptIntegerStringInLocale(trimmed, profile.decimalSeparator) && parsed >= -9223372036854775808.0 && parsed < 9223372036854775808.0 {
			return newVBScriptIntValue(int64(parsed), false), 0, true
		}
		return NewDouble(parsed), 0, true
	default:
		return NewEmpty(), vbscript.TypeMismatch, false
	}
}

// addValues performs numeric addition with integer and floating-point coercion.
// Returns Null if either operand is Null, following VBScript propagation rules.
// Classic VBScript semantics:
// - String + String => concatenation
// - Numeric + String => numeric addition, raising Type mismatch on invalid numeric string
// - Empty + Numeric => numeric addition with Empty as 0
// - Empty + String => string concatenation with Empty as ""
func (vm *VM) addValues(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(resolveCallable(vm, a))
	b = vm.unwrapADODBFieldValue(resolveCallable(vm, b))

	if isNull(a) || isNull(b) {
		return NewNull()
	}

	if isString(a) && isString(b) {
		return NewString(a.Str + b.Str)
	}

	if (isEmpty(a) && isString(b)) || (isString(a) && isEmpty(b)) {
		return NewString(vm.valueToString(a) + vm.valueToString(b))
	}

	if a.Type == VTDate || b.Type == VTDate {
		if sum, handled := vm.addOrSubtractDate(a, b, false); handled {
			return sum
		}
	}

	// VBScript numeric addition for String + Numeric only allows numeric types (VTInteger, VTCurrency, VTDouble).
	// Other types like VTBool or VTDate combined with VTString under '+' operator must raise Type mismatch.
	isStringANumericB := isString(a) && (b.Type == VTInteger || b.Type == VTCurrency || b.Type == VTDouble)
	isStringBNumericA := isString(b) && (a.Type == VTInteger || a.Type == VTCurrency || a.Type == VTDouble)
	bothNumeric := (isNumericLike(a) || isEmpty(a)) && (isNumericLike(b) || isEmpty(b))

	if isStringANumericB || isStringBNumericA || bothNumeric {
		left, leftOK := vm.coerceFloatStrict(a)
		right, rightOK := vm.coerceFloatStrict(b)
		if !leftOK || !rightOK {
			vm.raise(vbscript.TypeMismatch, "Type mismatch")
			return NewEmpty()
		}
		decSep := vm.localeProfile().decimalSeparator
		hasFloatStr := (isString(a) && !isVBScriptIntegerStringInLocale(strings.TrimSpace(a.Str), decSep)) ||
			(isString(b) && !isVBScriptIntegerStringInLocale(strings.TrimSpace(b.Str), decSep))
		if a.Type == VTDouble || b.Type == VTDouble || hasFloatStr {
			return NewDouble(left + right)
		}
		if a.Type == VTCurrency || b.Type == VTCurrency {
			if scaled, _, ok := floatToCurrencyScaled(left + right); ok && !isString(a) && !isString(b) {
				if a.Type == VTCurrency && b.Type == VTCurrency {
					return NewCurrency(a.Num + b.Num)
				}
				return NewCurrency(scaled)
			}
			return NewDouble(left + right)
		}
		return newVBScriptIntValue(int64(left+right), isLongSubtype(a) || isLongSubtype(b))
	}

	if isString(a) || isString(b) {
		vm.raise(vbscript.TypeMismatch, "Type mismatch")
		return NewEmpty()
	}

	if a.Type == VTDouble || b.Type == VTDouble {
		return NewDouble(vm.coerceFloat(a) + vm.coerceFloat(b))
	}
	if a.Type == VTCurrency || b.Type == VTCurrency {
		if scaled, _, ok := floatToCurrencyScaled(vm.coerceFloat(a) + vm.coerceFloat(b)); ok {
			return NewCurrency(scaled)
		}
		return NewDouble(vm.coerceFloat(a) + vm.coerceFloat(b))
	}
	return newVBScriptIntValue(vm.coerceInt64(a)+vm.coerceInt64(b), isLongSubtype(a) || isLongSubtype(b))
}

// addOrSubtractDate computes a + b, or a - b, when either operand is a Date, on the
// day serials as VBScript does. Measured on IIS (saborrie/iis-oracle runs 45592970246,
// 45595451489, 45595820636): the result is a Date, except that Date - Date is a Double
// and a serial outside the Date range stays a Double. The other operand may be a
// number, Boolean, Empty or numeric string (d + "1" is a Date); a non-numeric string
// raises Type mismatch. handled is false for operand types this does not cover.
func (vm *VM) addOrSubtractDate(a Value, b Value, subtract bool) (result Value, handled bool) {
	dateOperand := func(v Value) bool {
		return v.Type == VTDate || isNumericLike(v) || isEmpty(v) || isString(v)
	}
	if !dateOperand(a) || !dateOperand(b) {
		return Value{}, false
	}
	left, leftOK := vm.coerceFloatStrict(a)
	right, rightOK := vm.coerceFloatStrict(b)
	if !leftOK || !rightOK {
		vm.raise(vbscript.TypeMismatch, "Type mismatch")
		return NewEmpty(), true
	}
	if !subtract {
		return vbsDateOrDoubleFromSerial(left + right), true
	}
	if a.Type == VTDate && b.Type == VTDate {
		return NewDouble(left - right), true
	}
	return vbsDateOrDoubleFromSerial(left - right), true
}

// subtractValues performs numeric subtraction with integer and floating-point coercion.
// Returns Null if either operand is Null, following VBScript propagation rules.
func (vm *VM) subtractValues(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if isNull(a) || isNull(b) {
		return NewNull()
	}
	if a.Type == VTDate || b.Type == VTDate {
		if diff, handled := vm.addOrSubtractDate(a, b, true); handled {
			return diff
		}
	}
	if a.Type == VTDouble || b.Type == VTDouble {
		return NewDouble(vm.coerceFloat(a) - vm.coerceFloat(b))
	}
	if a.Type == VTCurrency || b.Type == VTCurrency {
		if a.Type == VTCurrency && b.Type == VTCurrency {
			return NewCurrency(a.Num - b.Num)
		}
		if scaled, _, ok := floatToCurrencyScaled(vm.coerceFloat(a) - vm.coerceFloat(b)); ok {
			return NewCurrency(scaled)
		}
		return NewDouble(vm.coerceFloat(a) - vm.coerceFloat(b))
	}
	return newVBScriptIntValue(vm.coerceInt64(a)-vm.coerceInt64(b), isLongSubtype(a) || isLongSubtype(b))
}

// multiplyValues performs numeric multiplication with integer and floating-point coercion.
// Returns Null if either operand is Null, following VBScript propagation rules.
func (vm *VM) multiplyValues(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if isNull(a) || isNull(b) {
		return NewNull()
	}
	// A Date multiplies as its day serial and gives a Double (IIS: d * 1 is 46388.4375).
	if a.Type == VTDouble || b.Type == VTDouble || a.Type == VTDate || b.Type == VTDate {
		return NewDouble(vm.coerceFloat(a) * vm.coerceFloat(b))
	}
	if a.Type == VTCurrency || b.Type == VTCurrency {
		if scaled, _, ok := floatToCurrencyScaled(vm.coerceFloat(a) * vm.coerceFloat(b)); ok {
			return NewCurrency(scaled)
		}
		return NewDouble(vm.coerceFloat(a) * vm.coerceFloat(b))
	}
	return newVBScriptIntValue(vm.coerceInt64(a)*vm.coerceInt64(b), isLongSubtype(a) || isLongSubtype(b))
}

// divideValues performs floating-point division and raises on division by zero.
// Returns Null if either operand is Null, following VBScript propagation rules.
func (vm *VM) divideValues(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if isNull(a) || isNull(b) {
		return NewNull()
	}
	divisor := vm.coerceFloat(b)
	if divisor == 0 {
		vm.raise(vbscript.DivisionByZero, "Division by zero")
		return NewEmpty()
	}
	return NewDouble(vm.coerceFloat(a) / divisor)
}

// intDivideValues performs integer division and raises on division by zero.
// Returns Null if either operand is Null, following VBScript propagation rules.
func (vm *VM) intDivideValues(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if isNull(a) || isNull(b) {
		return NewNull()
	}
	divisor := vm.coerceInt64(b)
	if divisor == 0 {
		vm.raise(vbscript.DivisionByZero, "Division by zero")
		return NewEmpty()
	}
	return NewInteger(vm.coerceInt64(a) / divisor)
}

// modValues performs modulo arithmetic using integer coercion.
// Returns Null if either operand is Null, following VBScript propagation rules.
func (vm *VM) modValues(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if isNull(a) || isNull(b) {
		return NewNull()
	}
	divisor := vm.coerceInt64(b)
	if divisor == 0 {
		vm.raise(vbscript.DivisionByZero, "Division by zero")
		return NewEmpty()
	}
	return NewInteger(vm.coerceInt64(a) % divisor)
}

// powValues performs exponentiation using floating-point coercion.
// Returns Null if either operand is Null, following VBScript propagation rules.
func (vm *VM) powValues(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if isNull(a) || isNull(b) {
		return NewNull()
	}
	return NewDouble(math.Pow(vm.coerceFloat(a), vm.coerceFloat(b)))
}

// compareValues compares two values for equality and relational operators.
func (vm *VM) compareValues(a Value, b Value) int {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if a.Type == VTNativeObject && a.Num == nativeObjectErr {
		a = vm.errPropertyValue("Number")
	}
	if b.Type == VTNativeObject && b.Num == nativeObjectErr {
		b = vm.errPropertyValue("Number")
	}
	if a.Type == VTNativeObject {
		if collectionValue, exists := vm.requestCollectionValueItems[a.Num]; exists {
			a = NewString(collectionValue.Joined())
		}
	}
	if b.Type == VTNativeObject {
		if collectionValue, exists := vm.requestCollectionValueItems[b.Num]; exists {
			b = NewString(collectionValue.Joined())
		}
	}

	if a.Type == VTNativeObject || b.Type == VTNativeObject || a.Type == VTObject || b.Type == VTObject || a.Type == VTBuiltin || b.Type == VTBuiltin {
		if a.Num < b.Num {
			return -1
		}
		if a.Num > b.Num {
			return 1
		}
		return 0
	}
	// A numeric string against a number compares numerically, as VBScript does
	// against a numeric literal ("131" > 82 is True, not a text comparison).
	if (a.Type == VTString) != (b.Type == VTString) {
		str, num := a, b
		if b.Type == VTString {
			str, num = b, a
		}
		if num.Type == VTInteger || num.Type == VTCurrency || num.Type == VTDouble {
			if parsed, _, ok := vbsParseNumericStringInLocale(str.Str, vm.localeProfile()); ok {
				left, right := parsed, vm.coerceFloat(num)
				if b.Type == VTString {
					left, right = right, left
				}
				if left < right {
					return -1
				}
				if left > right {
					return 1
				}
				return 0
			}
		}
	}
	if a.Type == VTString || b.Type == VTString {
		left := vm.valueToString(a)
		right := vm.valueToString(b)
		if vm.optionCompare == 1 {
			return vm.textCompare(left, right)
		}
		if left < right {
			return -1
		}
		if left > right {
			return 1
		}
		return 0
	}
	left := vm.coerceFloat(a)
	right := vm.coerceFloat(b)
	if left < right {
		return -1
	}
	if left > right {
		return 1
	}
	return 0
}

// logicalAnd performs boolean logical And or integer bitwise And.
// VBScript Null rule: Null And False = False; Null And True = Null; Null And Null = Null.
func (vm *VM) logicalAnd(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if isNull(a) || isNull(b) {
		// Short-circuit: Null And False -> False; otherwise Null
		if (!isNull(a) && !vm.asBool(a)) || (!isNull(b) && !vm.asBool(b)) {
			return NewBool(false)
		}
		return NewNull()
	}
	if a.Type == VTBool && b.Type == VTBool {
		return NewBool(vm.asBool(a) && vm.asBool(b))
	}
	return NewInteger(vm.coerceInt64(a) & vm.coerceInt64(b))
}

// logicalOr performs boolean logical Or or integer bitwise Or.
// VBScript Null rule: Null Or True = True; Null Or False = Null; Null Or Null = Null.
func (vm *VM) logicalOr(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if isNull(a) || isNull(b) {
		// Short-circuit: Null Or True -> True; otherwise Null
		if (!isNull(a) && vm.asBool(a)) || (!isNull(b) && vm.asBool(b)) {
			return NewBool(true)
		}
		return NewNull()
	}
	if a.Type == VTBool && b.Type == VTBool {
		return NewBool(vm.asBool(a) || vm.asBool(b))
	}
	return NewInteger(vm.coerceInt64(a) | vm.coerceInt64(b))
}

// logicalXor performs boolean logical Xor or integer bitwise Xor.
// Returns Null if either operand is Null.
func (vm *VM) logicalXor(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if isNull(a) || isNull(b) {
		return NewNull()
	}
	if a.Type == VTBool && b.Type == VTBool {
		return NewBool(vm.asBool(a) != vm.asBool(b))
	}
	return NewInteger(vm.coerceInt64(a) ^ vm.coerceInt64(b))
}

// logicalEqv performs boolean Eqv or integer bitwise equivalence.
// Returns Null if either operand is Null.
func (vm *VM) logicalEqv(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if isNull(a) || isNull(b) {
		return NewNull()
	}
	if a.Type == VTBool && b.Type == VTBool {
		return NewBool(vm.asBool(a) == vm.asBool(b))
	}
	return NewInteger(^(vm.coerceInt64(a) ^ vm.coerceInt64(b)))
}

// logicalImp performs boolean implication or integer bitwise implication.
// VBScript Null rule: False Imp Null = True; Null Imp True = True; else Null.
func (vm *VM) logicalImp(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if isNull(a) || isNull(b) {
		// False Imp Null = True (because !False = True, which dominates)
		if !isNull(a) && !vm.asBool(a) {
			return NewBool(true)
		}
		// Null Imp True = True (because True dominates)
		if !isNull(b) && vm.asBool(b) {
			return NewBool(true)
		}
		return NewNull()
	}
	if a.Type == VTBool && b.Type == VTBool {
		return NewBool(!vm.asBool(a) || vm.asBool(b))
	}
	return NewInteger((^vm.coerceInt64(a)) | vm.coerceInt64(b))
}

// concatValues concatenates two values as strings.
// VBScript '&' coerces Null operands to a zero-length string during concatenation.
// Uses vm.stringWorkBuffer as a reusable scratch buffer so that both parts can be
// appended without the hidden intermediate allocation that Go's '+' operator causes
// when one operand is not yet a string (e.g. numbers, dates, native objects).
func (vm *VM) concatValues(a Value, b Value) Value {
	a = vm.unwrapADODBFieldValue(a)
	b = vm.unwrapADODBFieldValue(b)
	if isNull(a) && isNull(b) {
		// Null & Null stays Null; only a lone Null operand coerces to "".
		// Measured on IIS: https://github.com/guimaraeslucas/axonasp/issues/188
		return NewNull()
	}
	if isNull(a) {
		a = NewString("")
	}
	if isNull(b) {
		b = NewString("")
	}
	// Fast-path: both operands are already plain strings — one allocation for the join.
	if a.Type == VTString && b.Type == VTString {
		vm.stringWorkBuffer = vm.stringWorkBuffer[:0]
		vm.stringWorkBuffer = append(vm.stringWorkBuffer, a.Str...)
		vm.stringWorkBuffer = append(vm.stringWorkBuffer, b.Str...)
		return NewString(string(vm.stringWorkBuffer))
	}
	// General path: convert both to strings via the reusable scratch buffer to avoid
	// creating a temporary intermediate string from the '+' operator.
	vm.stringWorkBuffer = vm.stringWorkBuffer[:0]
	vm.stringWorkBuffer = append(vm.stringWorkBuffer, vm.valueToString(a)...)
	vm.stringWorkBuffer = append(vm.stringWorkBuffer, vm.valueToString(b)...)
	return NewString(string(vm.stringWorkBuffer))
}
