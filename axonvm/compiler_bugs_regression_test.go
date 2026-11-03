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
	"errors"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"

	"g3pix.com.br/axonasp/v2/vbscript"
)

// repoFixturePath resolves one repository-relative fixture path from the axonvm package
// directory so regression tests never hard-code an absolute workspace location.
func repoFixturePath(t *testing.T, rel string) string {
	t.Helper()
	_, thisFile, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatalf("runtime.Caller failed to locate the test source file")
	}
	abs := filepath.Clean(filepath.Join(filepath.Dir(thisFile), "..", filepath.FromSlash(rel)))
	if _, err := os.Stat(abs); err != nil {
		t.Fatalf("fixture %s is not available: %v", abs, err)
	}
	return abs
}

// compileFixture compiles one ASP page fixture loaded from disk, anchored at its own
// path so relative "#include file=" directives resolve exactly like the server load path.
func compileFixture(t *testing.T, absPath string) *Compiler {
	t.Helper()
	content, err := os.ReadFile(absPath)
	if err != nil {
		t.Fatalf("failed reading %s: %v", absPath, err)
	}
	compiler := NewASPCompiler(string(content))
	compiler.SetSourceName(absPath)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("expected %s to compile cleanly, got: %v", filepath.Base(absPath), err)
	}
	return compiler
}

// TestOptionExplicitClassMethodScopeIsolation verifies that Option Explicit does not reject
// undeclared variables inside Class member bodies, matching Microsoft VBScript on IIS.
func TestOptionExplicitClassMethodScopeIsolation(t *testing.T) {
	compileFixture(t, repoFixturePath(t, "www/tests/bug1-repro.asp"))
}

// TestOptionExplicitStrictEnforcementOutsideClass verifies that Option Explicit still rejects
// undeclared variables in script scope and in module-level procedures outside Classes.
func TestOptionExplicitStrictEnforcementOutsideClass(t *testing.T) {
	testCases := []struct {
		name   string
		source string
	}{
		{
			name: "UndeclaredInStandaloneSub",
			source: `<% Option Explicit %>
<%
Sub TestSub()
	undeclaredVar = 42
End Sub
TestSub
%>`,
		},
		{
			name: "UndeclaredAtScriptRoot",
			source: `<% Option Explicit %>
<%
undeclaredGlobal = "invalid"
%>`,
		},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			compiler := NewASPCompiler(tc.source)
			err := compiler.Compile()
			if err == nil {
				t.Fatalf("expected a compile error for an undeclared variable outside a Class")
			}
			var syntaxErr *vbscript.VBSyntaxError
			if !errors.As(err, &syntaxErr) {
				t.Fatalf("expected VBScript syntax error, got %T (%v)", err, err)
			}
			if syntaxErr.Code != vbscript.VariableNotDefined {
				t.Fatalf("expected code %d (VariableNotDefined), got %d", vbscript.VariableNotDefined, syntaxErr.Code)
			}
		})
	}
}

// TestOptionExplicitClassMethodImplicitLocalIsolatedFromPageScope verifies that the implicit
// variable created inside a Class method stays procedure-local and is readable back through
// the method, while page scope never sees the name.
func TestOptionExplicitClassMethodImplicitLocalIsolatedFromPageScope(t *testing.T) {
	source := `<% Option Explicit %>
<%
Class ScopeProbe
	Function Probe()
		implicitLocal = "local-value"
		Probe = implicitLocal & "/" & TypeName(implicitLocal)
	End Function
End Class

Dim probe
Set probe = New ScopeProbe
Response.Write probe.Probe()
%>`

	if got := runVBSAndGetOutput(t, source); got != "local-value/String" {
		t.Fatalf("unexpected class-method output: got %q want %q", got, "local-value/String")
	}
}

// TestIndexedAssignmentInClassMethodCompiles verifies that indexed assignments inside Class
// methods are emitted as array writes and do not derail statement parsing.
func TestIndexedAssignmentInClassMethodCompiles(t *testing.T) {
	compileFixture(t, repoFixturePath(t, "www/tests/axonasp-bug2-repro/repro.asp"))
}

// TestIndexedAssignmentInClassMethodEmitsArraySet verifies that both a declared global
// dictionary and a local variable indexed assignment inside a Class method lower to OpArraySet
// instead of an implicit Me.<name>() member call.
func TestIndexedAssignmentInClassMethodEmitsArraySet(t *testing.T) {
	source := `<%
Dim dictEnvironment
Set dictEnvironment = Server.CreateObject("Scripting.Dictionary")

Class Probe
	Sub Run()
		Dim localDict
		Set localDict = Server.CreateObject("Scripting.Dictionary")
		dictEnvironment("KEY1") = 1
		dictEnvironment("KEY2") = "prefix_" & "suffix" & 123
		localDict("KEY3") = dictEnvironment("KEY1") + 1
	End Sub
End Class

Dim probe
Set probe = New Probe
probe.Run
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("failed compiling indexed assignments inside a Class method: %v", err)
	}

	bytecode := compiler.Bytecode()
	if got := countBytecodeOp(bytecode, OpArraySet); got != 3 {
		t.Fatalf("expected 3 OpArraySet writes for indexed assignments, got %d", got)
	}
	if got := countBytecodeOp(bytecode, OpMe); got != 0 {
		t.Fatalf("expected no OpMe statement-call hijack, got %d", got)
	}
}

// TestIndexedAssignmentInClassMethodRuntime verifies the end-to-end value written by an indexed
// assignment inside a Class method on a page-scope dictionary.
func TestIndexedAssignmentInClassMethodRuntime(t *testing.T) {
	source := `<%
Dim dictEnvironment
Set dictEnvironment = Server.CreateObject("Scripting.Dictionary")

Class Probe
	Sub Fill()
		dictEnvironment("COUNT") = 1 + 1
		dictEnvironment("LABEL") = "root" & "/" & "admin"
	End Sub
End Class

Dim probe
Set probe = New Probe
probe.Fill
Response.Write dictEnvironment("COUNT") & "|" & dictEnvironment("LABEL")
%>`

	if got := runVBSAndGetOutput(t, source); got != "2|root/admin" {
		t.Fatalf("unexpected indexed assignment output: got %q want %q", got, "2|root/admin")
	}
}

// TestGlobalSubCallInsideClassMethodStaysGlobal verifies that a module-level Sub invoked from a
// Class method is not rewritten into an implicit Me.<name>() member call.
func TestGlobalSubCallInsideClassMethodStaysGlobal(t *testing.T) {
	source := `<%
Dim result
result = ""

Sub GlobalHelper(value)
	result = result & value
End Sub

Class Probe
	Sub Run()
		GlobalHelper "A"
		GlobalHelper "B"
	End Sub
End Class

Dim probe
Set probe = New Probe
probe.Run
Response.Write result
%>`

	if got := runVBSAndGetOutput(t, source); got != "AB" {
		t.Fatalf("unexpected global helper output: got %q want %q", got, "AB")
	}
}

// TestLargeSubroutineCompilation verifies that large procedures compile without buffer limits,
// instruction-pointer drift, or excessive allocations.
func TestLargeSubroutineCompilation(t *testing.T) {
	var sb strings.Builder
	sb.WriteString("<%\nClass StressClass\nSub ProcessLargeData()\nDim x, arr(1000)\nx = 0\n")
	for i := 1; i <= 500; i++ {
		sb.WriteString("arr(1) = x + 1\n")
		sb.WriteString("x = x + 1\n")
	}
	sb.WriteString("End Sub\nEnd Class\n%>\n")

	compiler := NewASPCompiler(sb.String())
	if err := compiler.Compile(); err != nil {
		t.Fatalf("500-line subroutine compilation failed: %v", err)
	}
	if got := countBytecodeOp(compiler.Bytecode(), OpArraySet); got != 500 {
		t.Fatalf("expected 500 OpArraySet writes, got %d", got)
	}
}

// TestSetIndexedMemberCallAssignment covers `Set arr(i).Item(k) = obj` (aspJSON's loadJSON),
// which used to fail with "Expected member name after indexed call target in Set assignment".
func TestSetIndexedMemberCallAssignment(t *testing.T) {
	got := runVBSAndGetOutput(t, `<%
Dim level(2), d
Set level(0) = CreateObject("Scripting.Dictionary")
level(0).Add "k", ""
Set level(0).Item("k") = CreateObject("Scripting.Dictionary")
Set level(1) = level(0).Item("k")
level(1).Item("x") = "ok"
Response.Write TypeName(level(0).Item("k")) & ":" & level(0)("k")("x")
Set d = CreateObject("Scripting.Dictionary")
Set d.Item("a") = level(1)
Response.Write ":" & d("a")("x")
%>`)
	if got != "Dictionary:ok:ok" {
		t.Fatalf("got %q", got)
	}
}

// TestDoubleSlashLineComment covers legacy pages with a trailing `//note` after a statement,
// which IIS tolerates; it used to fail with "Unexpected token *vbscript.PunctuationToken".
func TestDoubleSlashLineComment(t *testing.T) {
	got := runVBSAndGetOutput(t, `<%
s = "a//b"//479698000000206355
Response.Write s & ":" & (6 / 3) //trailing
%>`)
	if got != "a//b:2" {
		t.Fatalf("got %q", got)
	}
}

// TestVersionedProgIDs covers versioned Windows ProgIDs (and WinHttp) that legacy pages create,
// which used to fail with 800401F3 "Invalid class string".
func TestVersionedProgIDs(t *testing.T) {
	got := runVBSAndGetOutput(t, `<%
For Each id In Array("WinHttp.WinHttpRequest.5.1", "MSXML2.ServerXMLHTTP.3.0", "MSXML2.ServerXMLHTTP.6.0", "MSXML2.XMLHTTP.3.0", "MSXML2.XMLHTTP.6.0", "MSXML2.DOMDocument.3.0", "MSXML2.DOMDocument.6.0")
	Set o = Server.CreateObject(id)
	Response.Write IsObject(o) & ","
Next
Set w = CreateObject("WinHttp.WinHttpRequest.5.1")
w.SetTimeouts 1000, 1000, 1000, 1000
w.Option(4) = &H3300
w.SetProxy 2, "proxy:8080"
Response.Write "ok"
%>`)
	if got != "True,True,True,True,True,True,True,ok" {
		t.Fatalf("got %q", got)
	}
}
