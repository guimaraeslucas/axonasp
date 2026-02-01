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

// VBArray holds array values alongside their lower bound to mimic VBScript arrays.
type VBArray struct {
	Lower  int
	Values []interface{}
}

// NewVBArray allocates a VBArray with the requested lower bound and size.
func NewVBArray(lower int, size int) *VBArray {
	if size < 0 {
		size = 0
	}
	return &VBArray{Lower: lower, Values: make([]interface{}, size)}
}

// NewVBArrayFromValues wraps an existing slice with an explicit lower bound.
func NewVBArrayFromValues(lower int, values []interface{}) *VBArray {
	if values == nil {
		values = []interface{}{}
	}
	return &VBArray{Lower: lower, Values: values}
}

// Upper returns the logical upper bound for the array.
func (a *VBArray) Upper() int {
	if len(a.Values) == 0 {
		return a.Lower - 1
	}
	return a.Lower + len(a.Values) - 1
}

// Len returns the number of stored elements.
func (a *VBArray) Len() int {
	return len(a.Values)
}

// Get retrieves a value using VBScript-style indexing.
func (a *VBArray) Get(index int) (interface{}, bool) {
	offset := index - a.Lower
	if offset < 0 || offset >= len(a.Values) {
		return nil, false
	}
	return a.Values[offset], true
}

// Set writes a value using VBScript-style indexing.
func (a *VBArray) Set(index int, value interface{}) bool {
	offset := index - a.Lower
	if offset < 0 || offset >= len(a.Values) {
		return false
	}
	a.Values[offset] = value
	return true
}

// toVBArray normalizes array-like values to VBArray while preserving backing storage.
func toVBArray(val interface{}) (*VBArray, bool) {
	switch v := val.(type) {
	case *VBArray:
		return v, true
	case []interface{}:
		return &VBArray{Lower: 0, Values: v}, true
	case [][]interface{}:
		// Convert 2D slice to VBArray of VBArrays
		values := make([]interface{}, len(v))
		for i, row := range v {
			values[i] = &VBArray{Lower: 0, Values: row}
		}
		return &VBArray{Lower: 0, Values: values}, true
	default:
		return nil, false
	}
}

// arrayBounds returns the lower and upper bounds for the given dimension (1-based).
func arrayBounds(val interface{}, dimension int) (int, int, bool) {
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
