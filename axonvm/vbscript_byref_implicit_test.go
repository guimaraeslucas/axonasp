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

// TestASPStaticObjectArgumentKeepsFallback verifies that the ByRef rewrite does not break
// the ASP StaticObjects lookup for <object runat="server"> identifiers: the first read of an
// undeclared page-scope identifier must still resolve the static object when it is used
// directly as a procedure argument.
func TestASPStaticObjectArgumentKeepsFallback(t *testing.T) {
	source := `<object runat="server" scope="Application" id="demoObj" progid="Scripting.FileSystemObject"></object>
<%
Sub show(v)
	Response.Write IsObject(v) & "|" & v.FileExists("__missing__.txt")
End Sub
show demoObj
%>`
	out := runVBSAndGetOutput(t, source)
	if out != "\nTrue|False" {
		t.Fatalf("static object argument lookup failed: got %q, want %q", out, "\nTrue|False")
	}
}

// TestVBScriptImplicitGlobalByRef covers implicitly declared page-scope variables
// passed to procedures. VBScript treats default parameters as ByRef even when the
// caller variable was never Dim'ed, so the callee must write back into the caller
// global slot for scalars, native objects and runtime class instances.
func TestVBScriptImplicitGlobalByRef(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{
			name: "BareCallScalar",
			source: `<%
Sub bump(x)
	x = x + 1
End Sub
bump z
bump z
Response.Write z
%>`,
			expected: "2",
		},
		{
			name: "ExpressionCallScalar",
			source: `<%
Function addOne(v)
	v = v + 1
	addOne = v
End Function
Response.Write addOne(z) & "|" & z
%>`,
			expected: "1|1",
		},
		{
			name: "NativeObjectWriteBack",
			source: `<%
Sub fill(target)
	Set target = Server.CreateObject("Scripting.Dictionary")
	target.Add "k", "v"
End Sub
fill d
Response.Write d("k")
%>`,
			expected: "v",
		},
		{
			name: "ClassInstanceWriteBack",
			source: `<%
Class Box
	Public Value
End Class
Sub makeBox(target)
	Set target = New Box
	target.Value = "boxed"
End Sub
makeBox b
Response.Write b.Value
%>`,
			expected: "boxed",
		},
		{
			name: "DeclaredGlobalStillByRef",
			source: `<%
Sub bump(x)
	x = x + 5
End Sub
Dim v
v = 1
bump v
Response.Write v
%>`,
			expected: "6",
		},
		{
			name: "ImplicitGlobalNativeObjectReassigned",
			source: `<%
Sub fill(target)
	Set target = Server.CreateObject("Scripting.Dictionary")
	target.Add "k", "v"
End Sub
fill d
fill d
Response.Write d("k") & "|" & d.Count
%>`,
			expected: "v|1",
		},
		{
			name: "ImplicitGlobalClassInstanceReassigned",
			source: `<%
Class Box
	Public Value
	Private Sub Class_Terminate()
		Response.Write "T" & Value & ";"
	End Sub
End Class
Sub make(target, v)
	Set target = New Box
	target.Value = v
End Sub
make b, "1"
make b, "2"
Response.Write b.Value
%>`,
			expected: "T1;2T2;",
		},
		{
			name: "LocalVariableObjectByRef",
			source: `<%
Class Box
	Public Value
End Class
Sub make(target, v)
	Set target = New Box
	target.Value = v
End Sub
Sub run()
	Dim local
	make local, "inner"
	Response.Write local.Value
End Sub
run
%>`,
			expected: "inner",
		},
		{
			name: "ByValParameterObjectNotWrittenBack",
			source: `<%
Sub make(ByVal target)
	Set target = Server.CreateObject("Scripting.Dictionary")
	target.Add "k", "v"
End Sub
make d
Response.Write IsEmpty(d)
%>`,
			expected: "True",
		},
		{
			name: "RecordsetWriteBack",
			source: `<%
Sub openRS(target)
	Set target = Server.CreateObject("ADODB.Recordset")
End Sub
openRS rs
Response.Write IsObject(rs)
%>`,
			expected: "True",
		},
		{
			name: "ForwardZeroArgFunctionArgumentStaysValue",
			source: `<%
Sub consume(x)
	Response.Write "got=" & x
End Sub
consume later
Function later()
	later = 7
End Function
%>`,
			expected: "got=7",
		},
		{
			name: "ForwardConstArgumentStaysValue",
			source: `<%
Sub consume(x)
	Response.Write "c=" & x
End Sub
consume LIMIT
Const LIMIT = 10
%>`,
			expected: "c=10",
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			out := runVBSAndGetOutput(t, tc.source)
			if out != tc.expected {
				t.Fatalf("test %q failed:\nexpected: %q\ngot:      %q", tc.name, tc.expected, out)
			}
		})
	}
}
