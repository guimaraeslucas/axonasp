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
	"math"
	"strconv"
	"strings"

	"g3pix.com.br/axonasp/vbscript"
)

// isObjectReferenceValue reports whether a value is a valid object reference for Is/Is Not.
func isObjectReferenceValue(v Value) bool {
	return v.Type == VTObject || v.Type == VTNativeObject
}

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
	case VTBool, VTInteger, VTDouble, VTDate:
		return true
	default:
		return false
	}
}

// coerceFloatStrict converts a value to float64 using VBScript-style numeric coercion.
// String conversion failure returns false so callers can raise Type mismatch.
func (vm *VM) coerceFloatStrict(v Value) (float64, bool) {
	v = resolveCallable(vm, v)
	switch v.Type {
	case VTEmpty:
		return 0, true
	case VTDouble:
		return v.Flt, true
	case VTBool, VTInteger, VTDate, VTNativeObject, VTBuiltin:
		return float64(v.Num), true
	case VTString:
		parsed, err := strconv.ParseFloat(strings.TrimSpace(v.Str), 64)
		if err != nil {
			return 0, false
		}
		return parsed, true
	default:
		return 0, false
	}
}

// coerceFloat converts a VM value into float64 using VBScript-like numeric coercion.
func (vm *VM) coerceFloat(v Value) float64 {
	v = resolveCallable(vm, v)
	switch v.Type {
	case VTDouble:
		return v.Flt
	case VTString:
		parsed, err := strconv.ParseFloat(strings.TrimSpace(v.Str), 64)
		if err == nil {
			return parsed
		}
		return 0
	default:
		return float64(v.Num)
	}
}

// coerceInt64 converts a VM value into int64 using VBScript-like numeric coercion.
// Float-to-integer conversions use Banker's Rounding (round-half-to-even) to match
// the behavior of VBScript's implicit type coercion for Integer and Long variables.
func (vm *VM) coerceInt64(v Value) int64 {
	v = resolveCallable(vm, v)
	switch v.Type {
	case VTBool, VTInteger, VTDate, VTNativeObject, VTBuiltin:
		return v.Num
	case VTDouble:
		return int64(math.RoundToEven(v.Flt))
	case VTString:
		parsed, err := strconv.ParseInt(strings.TrimSpace(v.Str), 10, 64)
		if err == nil {
			return parsed
		}
		parsedFloat, floatErr := strconv.ParseFloat(strings.TrimSpace(v.Str), 64)
		if floatErr == nil {
			return int64(math.RoundToEven(parsedFloat))
		}
		return 0
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
	v = resolveCallable(vm, v)
	switch v.Type {
	case VTBool, VTInteger, VTDate, VTNativeObject, VTBuiltin:
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
		if parsedFloat, err := strconv.ParseFloat(text, 64); err == nil {
			return int64(math.RoundToEven(parsedFloat)), true
		}
		return 0, false // non-numeric string — raises Type mismatch
	default:
		return 0, false
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
	a = resolveCallable(vm, a)
	b = resolveCallable(vm, b)

	if isNull(a) || isNull(b) {
		return NewNull()
	}

	if isString(a) && isString(b) {
		return NewString(a.Str + b.Str)
	}

	if (isEmpty(a) && isString(b)) || (isString(a) && isEmpty(b)) {
		return NewString(vm.valueToString(a) + vm.valueToString(b))
	}

	leftNumeric := isNumericLike(a) || isEmpty(a)
	rightNumeric := isNumericLike(b) || isEmpty(b)

	if (leftNumeric && (rightNumeric || isString(b))) || (rightNumeric && isString(a)) {
		left, leftOK := vm.coerceFloatStrict(a)
		right, rightOK := vm.coerceFloatStrict(b)
		if !leftOK || !rightOK {
			vm.raise(vbscript.TypeMismatch, "Type mismatch")
			return NewEmpty()
		}
		if a.Type == VTDouble || b.Type == VTDouble || (isString(a) && strings.Contains(strings.TrimSpace(a.Str), ".")) || (isString(b) && strings.Contains(strings.TrimSpace(b.Str), ".")) {
			return NewDouble(left + right)
		}
		return NewInteger(int64(left + right))
	}

	if isString(a) || isString(b) {
		vm.raise(vbscript.TypeMismatch, "Type mismatch")
		return NewEmpty()
	}

	if a.Type == VTDouble || b.Type == VTDouble {
		return NewDouble(vm.coerceFloat(a) + vm.coerceFloat(b))
	}
	return NewInteger(vm.coerceInt64(a) + vm.coerceInt64(b))
}

// subtractValues performs numeric subtraction with integer and floating-point coercion.
// Returns Null if either operand is Null, following VBScript propagation rules.
func (vm *VM) subtractValues(a Value, b Value) Value {
	if isNull(a) || isNull(b) {
		return NewNull()
	}
	if a.Type == VTDouble || b.Type == VTDouble {
		return NewDouble(vm.coerceFloat(a) - vm.coerceFloat(b))
	}
	return NewInteger(vm.coerceInt64(a) - vm.coerceInt64(b))
}

// multiplyValues performs numeric multiplication with integer and floating-point coercion.
// Returns Null if either operand is Null, following VBScript propagation rules.
func (vm *VM) multiplyValues(a Value, b Value) Value {
	if isNull(a) || isNull(b) {
		return NewNull()
	}
	if a.Type == VTDouble || b.Type == VTDouble {
		return NewDouble(vm.coerceFloat(a) * vm.coerceFloat(b))
	}
	return NewInteger(vm.coerceInt64(a) * vm.coerceInt64(b))
}

// divideValues performs floating-point division and raises on division by zero.
// Returns Null if either operand is Null, following VBScript propagation rules.
func (vm *VM) divideValues(a Value, b Value) Value {
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
	if isNull(a) || isNull(b) {
		return NewNull()
	}
	return NewDouble(math.Pow(vm.coerceFloat(a), vm.coerceFloat(b)))
}

// compareValues compares two values for equality and relational operators.
func (vm *VM) compareValues(a Value, b Value) int {
	if a.Type == VTNativeObject || b.Type == VTNativeObject || a.Type == VTObject || b.Type == VTObject || a.Type == VTBuiltin || b.Type == VTBuiltin {
		if a.Num < b.Num {
			return -1
		}
		if a.Num > b.Num {
			return 1
		}
		return 0
	}
	if a.Type == VTString || b.Type == VTString {
		left := vm.valueToString(a)
		right := vm.valueToString(b)
		if vm.optionCompare == 1 {
			left = strings.ToLower(left)
			right = strings.ToLower(right)
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
func (vm *VM) concatValues(a Value, b Value) Value {
	if isNull(a) {
		a = NewString("")
	}
	if isNull(b) {
		b = NewString("")
	}
	return NewString(vm.valueToString(a) + vm.valueToString(b))
}
