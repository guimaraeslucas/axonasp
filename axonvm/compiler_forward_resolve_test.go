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

import "testing"

// TestForwardConstInsideProcedure verifies that a global Const declared after the
// procedure body resolves to the constant value instead of an empty local slot.
func TestForwardConstInsideProcedure(t *testing.T) {
	source := `<%
Function getRate()
	getRate = RATE * 2
End Function

Sub showName()
	Response.Write APP_NAME & "|"
End Sub

showName
Response.Write getRate()

Const RATE = 21
Const APP_NAME = "Axon"
%>`
	if got := runVBSAndGetOutput(t, source); got != "Axon|42" {
		t.Fatalf("unexpected forward const output: got %q want %q", got, "Axon|42")
	}
}

// TestForwardConstInsideProcedureOptionExplicit verifies forward Const references
// inside procedures compile and resolve under Option Explicit.
func TestForwardConstInsideProcedureOptionExplicit(t *testing.T) {
	source := `<%
Option Explicit
Function getRate()
	getRate = RATE + 1
End Function
Response.Write getRate()
Const RATE = 9
%>`
	if got := runVBSAndGetOutput(t, source); got != "10" {
		t.Fatalf("unexpected forward const output: got %q want %q", got, "10")
	}
}

// TestForwardBareFunctionCallInExpression verifies that a parameterless Function
// referenced bare in an expression before its declaration is invoked.
func TestForwardBareFunctionCallInExpression(t *testing.T) {
	source := `<%
Dim x
x = myFunc + 1
Response.Write x & "|"
Response.Write myFunc & "|"
Response.Write CStr(myFunc * 2)
Function myFunc()
	myFunc = 41
End Function
%>`
	if got := runVBSAndGetOutput(t, source); got != "42|41|82" {
		t.Fatalf("unexpected forward bare function output: got %q want %q", got, "42|41|82")
	}
}

// TestForwardBareFunctionCallInsideProcedure verifies the same bare forward call
// from within another procedure body.
func TestForwardBareFunctionCallInsideProcedure(t *testing.T) {
	source := `<%
Function outer()
	outer = inner + 1
End Function
Response.Write outer()
Function inner()
	inner = 10
End Function
%>`
	if got := runVBSAndGetOutput(t, source); got != "11" {
		t.Fatalf("unexpected forward bare function output: got %q want %q", got, "11")
	}
}

// TestForwardBareFunctionCallOptionExplicit verifies bare forward function calls
// compile under Option Explicit.
func TestForwardBareFunctionCallOptionExplicit(t *testing.T) {
	source := `<%
Option Explicit
Dim x
x = myFunc + 1
Response.Write x
Function myFunc()
	myFunc = 1
End Function
%>`
	if got := runVBSAndGetOutput(t, source); got != "2" {
		t.Fatalf("unexpected forward bare function output: got %q want %q", got, "2")
	}
}

// TestForwardResolveUndefinedLocalStaysEmpty verifies that unknown identifiers
// inside procedures still behave as implicit empty locals without Option Explicit.
func TestForwardResolveUndefinedLocalStaysEmpty(t *testing.T) {
	source := `<%
Function probe()
	probe = "[" & undefinedThing & "]"
	undefinedThing = 5
End Function
Response.Write probe()
%>`
	if got := runVBSAndGetOutput(t, source); got != "[]" {
		t.Fatalf("unexpected undefined local output: got %q want %q", got, "[]")
	}
}
