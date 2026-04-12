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
	"strconv"
	"time"
)

type ValueType byte

const (
	VTEmpty ValueType = iota
	VTNull
	VTBool
	VTInteger
	VTDouble
	VTString
	VTDate
	VTArray
	VTObject
	VTNativeObject
	VTBuiltin
	VTUserSub
	// VTArgRef is a synthetic value type emitted only at call sites to carry a ByRef
	// slot reference alongside the argument value for post-call write-back.
	VTArgRef
)

type Value struct {
	Type  ValueType
	Num   int64   // Used for Bool (0/1), Integer, Date, NativeObject ID, and Builtin Index
	Flt   float64 // Used for Double
	Str   string  // Strings in Go are lightweight pointers
	Arr   *VBArray
	Names []string // Stores local names for VTUserSub or field names for VTObject
}

// String returns the string representation of the VBScript value.
func (v Value) String() string {
	switch v.Type {
	case VTEmpty:
		return ""
	case VTNull:
		return "Null"
	case VTBool:
		if v.Num != 0 {
			return "True"
		}
		return "False"
	case VTInteger:
		return fmt.Sprintf("%d", v.Num)
	case VTDouble:
		return fmt.Sprintf("%g", v.Flt)
	case VTString:
		return v.Str
	case VTArray:
		return "[Array]"
	case VTObject:
		if v.Num == 0 {
			return "" // Nothing has no string value
		}
		return fmt.Sprintf("[Object:%d]", v.Num)
	case VTNativeObject:
		return fmt.Sprintf("[NativeObject:%d]", v.Num)
	case VTBuiltin:
		return fmt.Sprintf("[Builtin:%d]", v.Num)
	case VTDate:
		return time.Unix(0, v.Num).UTC().Format(time.RFC3339)
	case VTUserSub:
		return "[UserSub]"
	case VTArgRef:
		return "[ArgRef]"
	default:
		return "Unknown"
	}
}

func NewInteger(v int64) Value {
	return Value{Type: VTInteger, Num: v}
}

func NewDouble(v float64) Value {
	return Value{Type: VTDouble, Flt: v}
}

func NewString(v string) Value {
	return Value{Type: VTString, Str: v}
}

func NewBool(v bool) Value {
	val := int64(0)
	if v {
		val = 1
	}
	return Value{Type: VTBool, Num: val}
}

// NewDate creates a VM date value from a Go time instance.
func NewDate(v time.Time) Value {
	return Value{Type: VTDate, Num: v.UTC().UnixNano()}
}

// NewNull returns a VM Null value. Any operation on Null propagates Null.
func NewNull() Value {
	return Value{Type: VTNull}
}

// NewEmpty returns a VM Empty value.
func NewEmpty() Value {
	return Value{Type: VTEmpty}
}

// NewUserSub creates a VM value for a user-defined Sub or Function.
// entryPoint is the bytecode offset where the body begins.
// paramCount is the number of declared parameters.
// localCount is the total number of local slots (params + locals + function return slot when applicable).
// isFunc is true for Function and false for Sub.
// byRefMask encodes which parameters are ByRef: bit i set = param i is ByRef.
// localNames is the list of symbol names for all local variable slots.
func NewUserSub(entryPoint int, paramCount int, localCount int, isFunc bool, byRefMask uint64, localNames []string) Value {
	if paramCount < 0 {
		paramCount = 0
	}
	if localCount < paramCount {
		localCount = paramCount
	}

	packed := (localCount << 12) | (paramCount & 0x0FFF)
	marker := float64(packed)
	if isFunc {
		marker += 0.5
	}
	v := Value{Type: VTUserSub, Num: int64(entryPoint), Flt: marker, Names: localNames}
	if byRefMask != 0 {
		v.Str = strconv.FormatUint(byRefMask, 10)
	}
	return v
}

// UserSubIsFunc reports whether a VTUserSub value represents a Function.
func (v Value) UserSubIsFunc() bool {
	return v.Flt != float64(int(v.Flt))
}

// UserSubParamCount returns the declared parameter count for a VTUserSub value.
func (v Value) UserSubParamCount() int {
	packed := int(v.Flt)
	if packed <= 0x0FFF {
		return packed
	}
	return packed & 0x0FFF
}

// UserSubLocalCount returns the total number of local slots for a VTUserSub value.
func (v Value) UserSubLocalCount() int {
	packed := int(v.Flt)
	if packed <= 0x0FFF {
		return packed
	}
	return packed >> 12

}

// UserSubByRefMask returns the ByRef parameter bitmask for a VTUserSub value.
// Bit i set means parameter i is declared ByRef.
func (v Value) UserSubByRefMask() uint64 {
	if v.Str == "" {
		return 0
	}
	mask, _ := strconv.ParseUint(v.Str, 10, 64)
	return mask
}

// ArgRefIsGlobal reports whether this VTArgRef references a global slot.
func (v Value) ArgRefIsGlobal() bool { return v.Flt == 1.0 }

// ArgRefIsLocal reports whether this VTArgRef references a local slot.
func (v Value) ArgRefIsLocal() bool { return v.Flt == 0.0 }

// ArgRefIsClassMember reports whether this VTArgRef references a class member slot.
func (v Value) ArgRefIsClassMember() bool { return v.Flt == 2.0 }

// ArgRefIdx returns the slot index encoded in a VTArgRef value.
func (v Value) ArgRefIdx() int { return int(v.Num) }
