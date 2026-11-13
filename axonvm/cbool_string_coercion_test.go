package axonvm

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 */

import "testing"

// TestCBoolAndIfConditionStringCoercionMatchIIS pins how a STRING coerces to a boolean,
// in the explicit CBool cast and in an If condition alike. Each line is measured on IIS
// (saborrie/iis-oracle): the CBool column in
// https://github.com/saborrie/iis-oracle/actions/runs/42522425088 (cbool-strings.vbs)
// and the If column in
// https://github.com/saborrie/iis-oracle/actions/runs/42511937887 (if-condition-coercion.asp).
//
// The line IIS draws: surrounding whitespace is forgiven for numeric strings (" 2 " is
// True) but not for the word booleans (" true " is Type mismatch, error 13).
//
// The rejected cases assert only that the Type mismatch is raised, not Err.Number's
// spelling: IIS reports the catalog code 13 where this engine currently reports the
// equivalent HRESULT -2146828275 (0x800A000D) — a separate normalization question
// that this matrix deliberately leaves open.
func TestCBoolAndIfConditionStringCoercionMatchIIS(t *testing.T) {
	src := `<%
Function DescribeCBool(ByVal s)
	Dim v
	On Error Resume Next
	v = CBool(s)
	If Err.Number <> 0 Then
		DescribeCBool = "ERR"
	Else
		DescribeCBool = CStr(v)
	End If
	On Error GoTo 0
End Function

Function DescribeIf(ByVal s)
	Dim taken
	On Error Resume Next
	If s Then
		taken = "True"
	Else
		taken = "False"
	End If
	If Err.Number <> 0 Then
		DescribeIf = "ERR"
	Else
		DescribeIf = taken
	End If
	On Error GoTo 0
End Function

Dim cases, i
cases = Array("0", "1", "-1", "0.0", " 2 ", "true", "False", " true ", "abc", "", "&h10", "&h0")
For i = 0 To UBound(cases)
	Response.Write DescribeCBool(cases(i)) & "/" & DescribeIf(cases(i)) & "|"
Next
%>`

	// One field per case, CBool result then If-branch result, in the order of `cases`.
	want := "False/False|" + // "0"
		"True/True|" + // "1"
		"True/True|" + // "-1"
		"False/False|" + // "0.0"
		"True/True|" + // " 2 " — numeric strings forgive padding
		"True/True|" + // "true"
		"False/False|" + // "False"
		"ERR/ERR|" + // " true " — word booleans do not
		"ERR/ERR|" + // "abc"
		"ERR/ERR|" + // ""
		"True/True|" + // "&h10"
		"False/False|" // "&h0"
	if got := runVBSAndGetOutput(t, src); got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}
