/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimaraes - G3pix Ltda
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

// VBArray stores VBScript-compatible array data and lower bound metadata.
type VBArray struct {
	Lower  int
	Values []Value
}

// GoVBArray stores a typed Go-facing array while keeping VBScript lower-bound information.
type GoVBArray struct {
	Lower  int
	Values []Value
}

// NewVBArray creates a new VBArray with the given lower bound and size.
func NewVBArray(lower int, size int) *VBArray {
	if size < 0 {
		size = 0
	}
	return &VBArray{Lower: lower, Values: make([]Value, size)}
}

// NewVBArrayFromValues creates a new VBArray from a value slice and lower bound.
func NewVBArrayFromValues(lower int, values []Value) *VBArray {
	if values == nil {
		values = []Value{}
	}
	return &VBArray{Lower: lower, Values: values}
}

// NewGoVBArray creates a GoVBArray with the given lower bound and size.
func NewGoVBArray(lower int, size int) *GoVBArray {
	if size < 0 {
		size = 0
	}
	return &GoVBArray{Lower: lower, Values: make([]Value, size)}
}

// NewGoVBArrayFromValues creates a GoVBArray from values and a lower bound.
func NewGoVBArrayFromValues(lower int, values []Value) *GoVBArray {
	if values == nil {
		values = []Value{}
	}
	return &GoVBArray{Lower: lower, Values: values}
}

// Upper returns the upper bound for the VBArray.
func (a *VBArray) Upper() int {
	if len(a.Values) == 0 {
		return a.Lower - 1
	}
	return a.Lower + len(a.Values) - 1
}

// Len returns the number of elements in the VBArray.
func (a *VBArray) Len() int {
	return len(a.Values)
}

// Get returns the value at the VBScript index and whether it exists.
func (a *VBArray) Get(index int) (Value, bool) {
	offset := index - a.Lower
	if offset < 0 || offset >= len(a.Values) {
		return Value{Type: VTEmpty}, false
	}
	return a.Values[offset], true
}

// Set writes a value at the VBScript index and returns true on success.
func (a *VBArray) Set(index int, value Value) bool {
	offset := index - a.Lower
	if offset < 0 || offset >= len(a.Values) {
		return false
	}
	a.Values[offset] = value
	return true
}

// toVBArray normalizes a VM value to a VBArray when possible.
func toVBArray(val Value) (*VBArray, bool) {
	if val.Type != VTArray || val.Arr == nil {
		return nil, false
	}
	return val.Arr, true
}

// arrayBounds returns lower and upper bounds for a 1-based dimension.
func arrayBounds(val Value, dimension int) (int, int, bool) {
	if dimension < 1 {
		return 0, 0, false
	}

	arr, ok := toVBArray(val)
	if !ok {
		return 0, 0, false
	}

	current := arr
	for d := 1; d < dimension; d++ {
		if len(current.Values) == 0 {
			return 0, -1, false
		}

		next, ok := toVBArray(current.Values[0])
		if !ok {
			return 0, 0, false
		}
		current = next
	}

	return current.Lower, current.Upper(), true
}

// ValueFromVBArray converts a VBArray into a VM array value.
func ValueFromVBArray(arr *VBArray) Value {
	if arr == nil {
		return Value{Type: VTNull}
	}
	return Value{Type: VTArray, Arr: arr}
}

// ValueFromGoArray converts a typed GoVBArray into a VM array value.
func ValueFromGoArray(arr *GoVBArray) Value {
	if arr == nil {
		return Value{Type: VTNull}
	}

	values := make([]Value, len(arr.Values))
	copy(values, arr.Values)
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(arr.Lower, values)}
}

// ValueFromValueSlice converts a Value slice into a zero-based VM array value.
func ValueFromValueSlice(values []Value) Value {
	cloned := make([]Value, len(values))
	copy(cloned, values)
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, cloned)}
}

// ToGoArray converts a VM array value into a typed GoVBArray.
func (v Value) ToGoArray() (*GoVBArray, bool) {
	if v.Type != VTArray || v.Arr == nil {
		return nil, false
	}

	values := make([]Value, len(v.Arr.Values))
	copy(values, v.Arr.Values)
	return &GoVBArray{Lower: v.Arr.Lower, Values: values}, true
}
