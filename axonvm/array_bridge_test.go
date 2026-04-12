package axonvm

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
import "testing"

// TestBuiltinArrayBounds verifies Array, IsArray, LBound, and UBound behavior for one-dimensional arrays.
func TestBuiltinArrayBounds(t *testing.T) {
	arr, err := vbsArray([]Value{NewInteger(10), NewString("x")})
	if err != nil {
		t.Fatalf("vbsArray returned error: %v", err)
	}
	if arr.Type != VTArray || arr.Arr == nil {
		t.Fatalf("expected VTArray with allocated payload, got: %#v", arr)
	}

	isArray, err := vbsIsArray([]Value{arr})
	if err != nil {
		t.Fatalf("vbsIsArray returned error: %v", err)
	}
	if isArray.Type != VTBool || isArray.Num != 1 {
		t.Fatalf("expected IsArray=True, got: %#v", isArray)
	}

	lb, err := vbsLBound([]Value{arr})
	if err != nil {
		t.Fatalf("vbsLBound returned error: %v", err)
	}
	if lb.Type != VTInteger || lb.Num != 0 {
		t.Fatalf("expected lower bound 0, got: %#v", lb)
	}

	ub, err := vbsUBound([]Value{arr})
	if err != nil {
		t.Fatalf("vbsUBound returned error: %v", err)
	}
	if ub.Type != VTInteger || ub.Num != 1 {
		t.Fatalf("expected upper bound 1, got: %#v", ub)
	}
}

// TestBuiltinNestedArrayBounds verifies second-dimension bounds by inspecting the first nested row.
func TestBuiltinNestedArrayBounds(t *testing.T) {
	row, _ := vbsArray([]Value{NewInteger(1), NewInteger(2), NewInteger(3)})
	table, _ := vbsArray([]Value{row})

	ub, err := vbsUBound([]Value{table, NewInteger(2)})
	if err != nil {
		t.Fatalf("vbsUBound returned error: %v", err)
	}
	if ub.Type != VTInteger || ub.Num != 2 {
		t.Fatalf("expected second-dimension upper bound 2, got: %#v", ub)
	}
}

// TestValueArrayRoundTrip verifies typed conversion between GoVBArray and VM array values.
func TestValueArrayRoundTrip(t *testing.T) {
	nested := ValueFromValueSlice([]Value{NewBool(true), NewDouble(3.5)})
	input := NewGoVBArrayFromValues(0, []Value{NewString("alpha"), NewInteger(7), nested})

	vmValue := ValueFromGoArray(input)
	if vmValue.Type != VTArray || vmValue.Arr == nil {
		t.Fatalf("expected VTArray from GoVBArray, got: %#v", vmValue)
	}

	array, ok := vmValue.ToGoArray()
	if !ok {
		t.Fatalf("expected ToGoArray to succeed")
	}

	if array.Lower != 0 || len(array.Values) != 3 {
		t.Fatalf("unexpected top-level GoVBArray shape: %#v", array)
	}

	nestedValue := array.Values[2]
	nestedArray, ok := nestedValue.ToGoArray()
	if !ok {
		t.Fatalf("expected nested array value, got: %#v", nestedValue)
	}
	if nestedArray.Lower != 0 || len(nestedArray.Values) != 2 {
		t.Fatalf("unexpected nested GoVBArray shape: %#v", nestedArray)
	}
}
