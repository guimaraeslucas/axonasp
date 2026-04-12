/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
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
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"

	"g3pix.com.br/axonasp/axonvm/asp"
	"g3pix.com.br/axonasp/vbscript"
)

// TestASPCompileSupportsBooleanLiteralArguments verifies boolean literals compile in statement-style member calls.
func TestASPCompileSupportsBooleanLiteralArguments(t *testing.T) {
	source := `<%
Response.Buffer True
Response.Write "ok"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
}

// TestASPCompileClassDeclarationDoesNotBreakTopLevel verifies class blocks compile and top-level script still runs.
func TestASPCompileClassDeclarationDoesNotBreakTopLevel(t *testing.T) {
	source := `<%
Class Widget
End Class
Response.Write "ok"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected ok output, got %q", output.String())
	}
}

// TestASPChainedPropertyAssignmentThroughIntermediateObject verifies that
// statement assignments like h.Obj.Value = 42 resolve intermediate members
// before setting the final property, matching IIS VBScript behavior.
func TestASPChainedPropertyAssignmentThroughIntermediateObject(t *testing.T) {
	source := `<%
Class Inner
	Public Value
End Class

Class Holder
	Private pObj
	Private Sub Class_Initialize()
		Set pObj = New Inner
	End Sub
	Public Function Obj()
		Set Obj = pObj
	End Function
End Class

Dim h
Set h = New Holder
h.Obj.Value = 42
Response.Write h.Obj.Value
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "42" {
		t.Fatalf("expected 42 output, got %q", output.String())
	}
}

// TestASPChainedPropertyAssignmentThroughFunctionObject verifies chained member
// sets when the intermediate member is one function that returns an object.
func TestASPChainedPropertyAssignmentThroughFunctionObject(t *testing.T) {
	source := `<%
Class JsonContainer
	Public recordsetPaging
End Class

Class AspLite
	Private pJson
	Private Sub Class_Initialize()
		Set pJson = Nothing
	End Sub
	Public Function json
		If pJson Is Nothing Then
			Set pJson = New JsonContainer
		End If
		Set json = pJson
	End Function
End Class

Dim aspl
Set aspl = New AspLite
aspl.json.recordsetPaging = True
Response.Write CStr(aspl.json.recordsetPaging)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "True" {
		t.Fatalf("expected True output, got %q", output.String())
	}
}

// TestASPClassPropertyExitPropertyKeywordLike verifies Exit Property inside a
// class Property Get compiles and executes when Property is parsed as a
// keyword-like identifier token.
func TestASPClassPropertyExitPropertyKeywordLike(t *testing.T) {
	source := `<%
Class LoginBox
	Private p1
	Private p2
	Public Property Get CurrentPW
		If p1 <> "" Then
			CurrentPW = p1
			Exit Property
		End If
		CurrentPW = p2
		Exit Property
	End Property
	Public Sub SeedValues()
		p1 = ""
		p2 = "fallback"
	End Sub
End Class

Dim l
Set l = New LoginBox
l.SeedValues
Response.Write l.CurrentPW
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "fallback" {
		t.Fatalf("expected fallback output, got %q", output.String())
	}
}

// TestASPClassMethodSessionCollectionAssignmentCompile verifies Session(...)=... in
// class method scope stays bound to the global ASP Session object rather than an
// implicit Me.<member> statement-call rewrite.
func TestASPClassMethodSessionCollectionAssignmentCompile(t *testing.T) {
	source := `<%
Class Contact
	Public iId
End Class

Class LoginBox
	Private p_contact
	Private Sub Class_Initialize
		Set p_contact = New Contact
		p_contact.iId = 7
	End Sub
	Public Function getUFP
		Session("userfolderpath")="userfiles/" & p_contact.iId & "/"
		getUFP = "ok"
	End Function
End Class
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
}

// TestASPLargeScriptJumpTargets verifies absolute jump targets can exceed the
// previous 64K bytecode ceiling without breaking compilation or execution.
func TestASPLargeScriptJumpTargets(t *testing.T) {
	var source strings.Builder
	source.Grow(320000)
	source.WriteString("<%\nIf False Then\n")
	for i := 0; i < 9000; i++ {
		source.WriteString("Response.Write \"xxxxxxxxxx\"\n")
	}
	source.WriteString("End If\nResponse.Write \"ok\"\n%>")

	compiler := NewASPCompiler(source.String())
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected ok output, got %q", output.String())
	}
}

// TestASPPageScopePrivateVariableDeclaration verifies Classic ASP page-scope
// `Private name` declarations compile and behave like module-level variables.
func TestASPPageScopePrivateVariableDeclaration(t *testing.T) {
	source := `<%
Private pageCounter
pageCounter = 1
Response.Write pageCounter
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "1" {
		t.Fatalf("expected 1 output, got %q", output.String())
	}
}

// TestASPClassExpressionCallBindsForwardGlobalFunction verifies that a bare
// function call used inside a class expression keeps global resolution when the
// target is a later page-level function, matching IIS/VBScript behavior.
func TestASPClassExpressionCallBindsForwardGlobalFunction(t *testing.T) {
	source := `<%
Class LoginBox
	Public Function ReadSession(key)
		ReadSession = ConvertBool(Session(key))
	End Function
End Class

Function ConvertBool(value)
	ConvertBool = CStr(CBool(value))
End Function

Dim l
Set l = New LoginBox
Session("flag") = True
Response.Write l.ReadSession("flag")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "True" {
		t.Fatalf("expected True output, got %q", output.String())
	}
}

// TestASPClassExpressionCallBindsForwardClassMethod verifies unqualified calls
// inside class methods can resolve to later class members when no global symbol exists.
func TestASPClassExpressionCallBindsForwardClassMethod(t *testing.T) {
	source := `<%
Class Wrapper
	Public Function A()
		A = B()
	End Function
	Public Function B()
		B = "ok"
	End Function
End Class

Dim w
Set w = New Wrapper
Response.Write w.A()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected ok output, got %q", output.String())
	}
}

// TestASPHoistingClassSubFunctionDeclarations verifies IIS-like pre-binding for Class/Sub/Function.
func TestASPHoistingClassSubFunctionDeclarations(t *testing.T) {
	source := `<%
Dim obj
Set obj = New LateClass

Call LateSub()
Response.Write LateFunc()
Response.Write obj.Name()

Class LateClass
	Public Function Name()
		Name = "class"
	End Function
End Class

Sub LateSub()
	Response.Write "sub"
End Sub

Function LateFunc()
	LateFunc = "func"
End Function
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "subfuncclass" {
		t.Fatalf("expected subfuncclass output, got %q", output.String())
	}
}

// TestASPCompileRejectsNestedClassDeclaration verifies nested Class blocks are rejected.
// Classic ASP/VBScript does not support class declarations inside another class body.
func TestASPCompileRejectsNestedClassDeclaration(t *testing.T) {
	source := `<%
Class Outer
    Class Inner
    End Class
End Class
%>`

	compiler := NewASPCompiler(source)
	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compile error for nested class declaration, got nil")
	}

	var vbErr *vbscript.VBSyntaxError
	if !errors.As(err, &vbErr) {
		t.Fatalf("expected VBError, got %T: %v", err, err)
	}
	if vbErr.Code != vbscript.SyntaxError {
		t.Fatalf("expected SyntaxError, got code %d (%s)", vbErr.Code, vbErr.Description)
	}
}

// TestASPIfElseIfBlockExecutesExpectedBranch verifies block If chains execute ElseIf branches.
func TestASPIfElseIfBlockExecutesExpectedBranch(t *testing.T) {
	source := `<%
Dim n
n = 2

If n = 1 Then
	Response.Write "one"
ElseIf n = 2 Then
	Response.Write "two"
Else
	Response.Write "other"
End If
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "two" {
		t.Fatalf("expected two output, got %q", output.String())
	}
}

// TestASPIfElseIfInlineExecutesExpectedBranch verifies one-line If chains execute ElseIf branches.
func TestASPIfElseIfInlineExecutesExpectedBranch(t *testing.T) {
	source := `<% Dim n: n = 3: If n = 1 Then Response.Write "one" ElseIf n = 3 Then Response.Write "three" Else Response.Write "other" %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "three" {
		t.Fatalf("expected three output, got %q", output.String())
	}
}

// TestASPInlineIfWithExplicitEndIf verifies single-line If accepts trailing End If.
func TestASPInlineIfWithExplicitEndIf(t *testing.T) {
	source := `<% Dim s: s = "abc?x=1": If InStr(s, "?") > 0 Then s = Left(s, InStr(s, "?") - 1) : End If : Response.Write s %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "abc" {
		t.Fatalf("expected abc output, got %q", output.String())
	}
}

// TestASPInlineIfElseWithExplicitEndIf verifies one-line If/Else with trailing End If.
func TestASPInlineIfElseWithExplicitEndIf(t *testing.T) {
	source := `<% Dim n: n = 0: If n > 0 Then Response.Write "yes" Else Response.Write "no" : End If %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "no" {
		t.Fatalf("expected no output, got %q", output.String())
	}
}

// TestASPInlineIfElseIfWithExplicitEndIf verifies one-line If/ElseIf with trailing End If.
func TestASPInlineIfElseIfWithExplicitEndIf(t *testing.T) {
	source := `<% Dim n: n = 3: If n = 1 Then Response.Write "one" ElseIf n = 3 Then Response.Write "three" Else Response.Write "other" : End If %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "three" {
		t.Fatalf("expected three output, got %q", output.String())
	}
}

// TestASPColonStatementSeparatorExecutesAllStatements verifies colon-separated inline statements execute in order.
func TestASPColonStatementSeparatorExecutesAllStatements(t *testing.T) {
	source := `<% Dim s: s = "": s = s & "A": s = s & "B": If Len(s) = 2 Then s = s & "C": Response.Write s %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ABC" {
		t.Fatalf("expected ABC output, got %q", output.String())
	}
}

// TestASPSingleLineIfThenInsideFunctionWithEndFunction verifies that a single-line
// "If condition Then statement" inside a public/private function body compiles and
// runs correctly when "End Function" appears on the following line.
// Regression: the compatibility block for optional trailing "End If" was
// incorrectly skipping line terminators and consuming the "End" token that
// belonged to "End Function", causing a spurious "Missing Identifier" error.
func TestASPSingleLineIfThenInsideFunctionWithEndFunction(t *testing.T) {
	source := `<%
Public Function GetFileExtension(str)
    Dim Pos : Pos = InStrRev(str, ".")
    If Pos > 0 Then GetFileExtension = Mid(str, Pos + 1)
End Function
Response.Write GetFileExtension("document.pdf")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "pdf" {
		t.Fatalf("expected pdf, got %q", output.String())
	}
}

// TestASPSingleLineIfThenInsideSubWithEndSub verifies the same inline If fix for Sub bodies.
func TestASPSingleLineIfThenInsideSubWithEndSub(t *testing.T) {
	source := `<%
Dim result
Sub CheckPositive(n)
    If n > 0 Then result = "pos"
End Sub
CheckPositive 5
Response.Write result
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "pos" {
		t.Fatalf("expected pos, got %q", output.String())
	}
}

// TestASPLineContinuationUnderscore verifies VBScript line continuation with underscore.
func TestASPLineContinuationUnderscore(t *testing.T) {
	source := `<%
Dim s
s = "Ax" & _
    "on" & _
    "ASP"
Response.Write s
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "AxonASP" {
		t.Fatalf("expected AxonASP output, got %q", output.String())
	}
}

// TestASPWhileWEndCompilesAndRuns verifies While...WEnd loops compile and execute.
func TestASPWhileWEndCompilesAndRuns(t *testing.T) {
	source := `<%
Dim i
i = 0
While i < 3
	i = i + 1
Wend
Response.Write i
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "3" {
		t.Fatalf("expected 3 output, got %q", output.String())
	}
}

// TestASPWhileWEndSupportsAndTypeNamePattern verifies While conditions like jsonObject.class.asp compile.
func TestASPWhileWEndSupportsAndTypeNamePattern(t *testing.T) {
	source := `<%
Dim tmpObj
Set tmpObj = Nothing

While IsObject(tmpObj) And TypeName(tmpObj) = "JSONobject"
	Response.Write "x"
Wend

Response.Write "ok"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected ok output, got %q", output.String())
	}
}

// TestASPCompileAllowsOmittedArgInNoParenMemberCall verifies VBScript omitted args in member calls without parentheses.
func TestASPCompileAllowsOmittedArgInNoParenMemberCall(t *testing.T) {
	source := `<%
Dim rs
Set rs = CreateObject("ADODB.Recordset")
rs.Fields.Append "ID", adInteger, , adFldKeyColumn
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
}

// TestASPCompileAllowsOmittedArgInParenCall verifies VBScript omitted args in parenthesized function calls.
func TestASPCompileAllowsOmittedArgInParenCall(t *testing.T) {
	source := `<%
Function F(a, b, c)
	F = IsEmpty(b)
End Function

If F(1, , 3) Then
	Response.Write "ok"
End If
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected ok output, got %q", output.String())
	}
}

// TestASPADODBRecordsetDisconnectedOpenAndFieldAssignment verifies disconnected in-memory recordset usage compatible with Classic ASP.
func TestASPADODBRecordsetDisconnectedOpenAndFieldAssignment(t *testing.T) {
	source := `<%
Dim rs
Set rs = CreateObject("ADODB.Recordset")

rs.CursorType = adOpenKeyset
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic

rs.Fields.Append "ID", adInteger, , adFldKeyColumn
rs.Fields.Append "Nome", adVarChar, 50, adFldMayBeNull
rs.Fields.Append "Valor", adDecimal, 14, adFldMayBeNull
rs.Fields("Valor").NumericScale = 2

rs.Open
rs.AddNew
rs("ID") = 1
rs("Nome") = "Nome 1"
rs("Valor") = 10.99
rs.Update

rs.MoveFirst
Response.Write rs("Nome") & "|" & rs("Valor")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Nome 1|10.99" {
		t.Fatalf("expected Nome 1|10.99 output, got %q", output.String())
	}
}

// TestASPSetCreateObjectBuiltinReturnsObject verifies Set var = CreateObject(...) stores a native object reference.
func TestASPSetCreateObjectBuiltinReturnsObject(t *testing.T) {
	source := `<%
Dim d
Set d = CreateObject("Scripting.Dictionary")
Response.Write CStr(IsObject(d)) & "|"
d.Add "k", "v"
Response.Write d.Count
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "True|1" {
		t.Fatalf("unexpected Set/CreateObject output: %q", output.String())
	}
}

// TestASPADODBRecordsetFieldsForEach verifies For Each over rs.Fields returns field proxies with Name/Value.
func TestASPADODBRecordsetFieldsForEach(t *testing.T) {
	source := `<%
Dim rs, f
Set rs = CreateObject("ADODB.Recordset")

rs.Fields.Append "ID", adInteger, , adFldKeyColumn
rs.Fields.Append "Nome", adVarChar, 50, adFldMayBeNull
rs.Open
rs.AddNew
rs("ID") = 7
rs("Nome") = "Item"
rs.Update

rs.MoveFirst
For Each f In rs.Fields
	Response.Write f.Name & ":" & f.Value & "|"
Next
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ID:7|Nome:Item|" {
		t.Fatalf("expected ID:7|Nome:Item| output, got %q", output.String())
	}
}

// TestASPADODBRecordsetOpenTableName verifies Recordset.Open accepts a bare table name.
func TestASPADODBRecordsetOpenTableName(t *testing.T) {
	source := `<%
Dim conn, rs, dbPath
dbPath = Server.MapPath("/recordset-open-table.db")

Set conn = CreateObject("ADODB.Connection")
conn.Open "sqlite:" & dbPath
conn.Execute "CREATE TABLE sample_items (id INTEGER, name TEXT)"
conn.Execute "INSERT INTO sample_items (id, name) VALUES (1, 'alpha')"

Set rs = CreateObject("ADODB.Recordset")
rs.Open "sample_items", conn
Response.Write rs("name")
rs.Close
conn.Close
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	host.Server().SetRequestPath("/tests/test_adodb_open.asp")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "alpha" {
		t.Fatalf("expected alpha output, got %q", output.String())
	}
}

// TestASPADODBFieldProxyDefaultValueCoercion verifies native ADODB.Field proxies coerce
// to their default Value property in value contexts (e.g., Response.Write).
func TestASPADODBFieldProxyDefaultValueCoercion(t *testing.T) {
	source := `<%
Dim rs, f
Set rs = CreateObject("ADODB.Recordset")

rs.Fields.Append "Nome", adVarChar, 50, adFldMayBeNull
rs.Open
rs.AddNew
rs("Nome") = "Item"
rs.Update
rs.MoveFirst

Set f = rs.Fields("Nome")
Response.Write f
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Item" {
		t.Fatalf("expected Item output, got %q", output.String())
	}
}

// TestASPADODBFieldsItemValueContextCoercion verifies chained Fields.Item(...)
// expressions coerce to ADODB.Field.Value in concatenation/output contexts.
func TestASPADODBFieldsItemValueContextCoercion(t *testing.T) {
	source := `<%
Dim rs
Set rs = CreateObject("ADODB.Recordset")

rs.Fields.Append "Nome", adVarChar, 50, adFldMayBeNull
rs.Open
rs.AddNew
rs("Nome") = "Item"
rs.Update
rs.MoveFirst

Response.Write "Name=" & rs.Fields.Item("Nome")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Name=Item" {
		t.Fatalf("expected Name=Item output, got %q", output.String())
	}
}

// TestASPNativeObjectArgumentPassThrough verifies native objects passed as Sub
// arguments are not coerced through __default__ before call dispatch.
func TestASPNativeObjectArgumentPassThrough(t *testing.T) {
	source := `<%
Sub EchoType(obj)
	Response.Write TypeName(obj)
End Sub

Dim d
Set d = CreateObject("Scripting.Dictionary")
Call EchoType(d)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Dictionary" {
		t.Fatalf("expected Dictionary output, got %q", output.String())
	}
}

// TestASPStatementChainedMemberCall verifies statement-style chained calls resolve
// intermediate members before invoking the final method.
func TestASPStatementChainedMemberCall(t *testing.T) {
	source := `<%
Class JsonSink
	Public Sub Dump(v)
		Response.Write TypeName(v)
	End Sub
End Class

Class Root
	Public Function Json
		Set Json = New JsonSink
	End Function
End Class

Dim r, d
Set r = New Root
Set d = CreateObject("Scripting.Dictionary")
r.Json.Dump(d)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Dictionary" {
		t.Fatalf("expected Dictionary output, got %q", output.String())
	}
}

// TestASPStatementMemberCallUsesMemberCallResultArgument verifies statement-style
// member calls preserve nested member-call return values as arguments.
func TestASPStatementMemberCallUsesMemberCallResultArgument(t *testing.T) {
	source := `<%
Class Reader
	Public Function Read()
		Read = "payload"
	End Function
End Class

Class Writer
	Public Sub Write(value)
		Response.Write value
	End Sub
End Class

Dim readerObj, writerObj
Set readerObj = New Reader
Set writerObj = New Writer
writerObj.Write readerObj.Read()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "payload" {
		t.Fatalf("expected payload output, got %q", output.String())
	}
}

// TestASPNestedResumeNextErrorDoesNotSkipCallerStatement verifies a callee's
// absorbed mid-function error does not skip the caller's remaining statement.
func TestASPNestedResumeNextErrorDoesNotSkipCallerStatement(t *testing.T) {
	source := `<%
Class Reader
	Public Function Read()
		On Error Resume Next
		Dim d
		Set d = CreateObject("Scripting.Dictionary")
		Read = "payload"
		d.Nope
	End Function
End Class

Class Writer
	Public Sub Write(value)
		Response.Write "[" & value & "]"
	End Sub
End Class

Dim readerObj, writerObj, noop
Set readerObj = New Reader
Set writerObj = New Writer
writerObj.Write readerObj.Read()
noop = 1
Response.Write "|after|"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "[payload]|after|" {
		t.Fatalf("expected [payload]|after| output, got %q", output.String())
	}
}

// TestASPUBoundInvalidDimensionRaisesError verifies UBound invalid dimension sets Err under On Error Resume Next.
func TestASPUBoundInvalidDimensionRaisesError(t *testing.T) {
	source := `<%
On Error Resume Next
Dim a
a = Array(1,2,3)
UBound a, 2
If Err.Number <> 0 Then
	Response.Write "ok"
Else
	Response.Write "fail"
End If
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected ok output, got %q", output.String())
	}
}

// TestASPNumDimensionsPatternTerminates verifies the classic NumDimensions probing loop terminates.
func TestASPNumDimensionsPatternTerminates(t *testing.T) {
	source := `<%
Function NumDimensions(ByRef arr)
	Dim dimensions
	dimensions = 0
	On Error Resume Next
	Do While Err.number = 0
		dimensions = dimensions + 1
		UBound arr, dimensions
	Loop
	On Error Goto 0
	NumDimensions = dimensions - 1
End Function

Dim a
a = Array(1,2,3)
Response.Write NumDimensions(a)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "1" {
		t.Fatalf("expected 1 output, got %q", output.String())
	}
}

// TestASPClassMemberGetInvokesZeroArgSub verifies member-get expressions can invoke zero-arg class Sub methods.
func TestASPClassMemberGetInvokesZeroArgSub(t *testing.T) {
	source := `<%
Class Writer
	Public Sub Write
		Response.Write "ok"
	End Sub
End Class

Dim w
Set w = New Writer
Response.Write w.Write
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected ok output, got %q", output.String())
	}
}

// TestASPClassTerminateCanAccessProperty verifies Class_Terminate can read and clear a property on the same instance.
func TestASPClassTerminateCanAccessProperty(t *testing.T) {
	source := `<%
Class Holder
	Private i_value

	Public Property Get value
		Set value = i_value
	End Property

	Public Property Set value(v)
		Set i_value = v
	End Property

	Private Sub class_terminate
		If IsObject(value) Then Set value = Nothing
	End Sub
End Class

Class Child
End Class

Dim h
Set h = New Holder
Set h.value = New Child
Set h = Nothing
Response.Write "ok"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected ok output, got %q", output.String())
	}
}

// TestASPClassImplicitMemberCallWithArgs verifies unqualified class method calls with arguments compile and execute correctly.
func TestASPClassImplicitMemberCallWithArgs(t *testing.T) {
	source := `<%
Class Calc
	Public Function Add(a, b)
		Add = a + b
	End Function

	Public Function Twice(v)
		Twice = Add(v, v)
	End Function
End Class

Dim c
Set c = New Calc
Response.Write c.Twice(5)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "10" {
		t.Fatalf("expected 10 output, got %q", output.String())
	}
}

// TestASPClassCanCallPrivateMemberInternally verifies same-instance access to private methods remains allowed.
func TestASPClassCanCallPrivateMemberInternally(t *testing.T) {
	source := `<%
Class Worker
	Private Function Inc(v)
		Inc = v + 1
	End Function

	Public Function Run(v)
		Run = Inc(v)
	End Function
End Class

Dim w
Set w = New Worker
Response.Write w.Run(9)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "10" {
		t.Fatalf("expected 10 output, got %q", output.String())
	}
}

// TestASPCompileClassWithMembersKeepsTopLevelExecution verifies class member declarations are safely consumed in compile flow.
func TestASPCompileClassWithMembersKeepsTopLevelExecution(t *testing.T) {
	source := `<%
Class Counter
	Private value
	Public Sub SetValue(v)
		value = v
	End Sub
	Public Function GetValue()
		GetValue = value
	End Function
End Class
Response.Write "ok"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected ok output, got %q", output.String())
	}
}

// TestASPTypeNameReturnsRuntimeClassName verifies TypeName returns class names for VTObject instances.
func TestASPTypeNameReturnsRuntimeClassName(t *testing.T) {
	source := `<%
Class JSONobject
End Class

Dim obj
Set obj = New JSONobject
Response.Write TypeName(obj)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "JSONobject" {
		t.Fatalf("expected JSONobject output, got %q", output.String())
	}
}

// TestASPEvalSupportsJSONSpecialValues verifies Eval supports JSON parser literals.
func TestASPEvalSupportsJSONSpecialValues(t *testing.T) {
	source := `<%
Response.Write TypeName(Eval("true")) & "|"
Response.Write TypeName(Eval("false")) & "|"
Response.Write TypeName(Eval("null")) & "|"
Response.Write TypeName(Eval("undefined"))
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Boolean|Boolean|Null|Empty" {
		t.Fatalf("unexpected Eval output: %q", output.String())
	}
}

// TestASPEvalDoesNotRunClassTerminate verifies Eval does not trigger object termination during the active request.
func TestASPEvalDoesNotRunClassTerminate(t *testing.T) {
	source := `<%
Class Probe
	Private arr

	Private Sub Class_Initialize
		ReDim arr(0)
		arr(0) = "ok"
	End Sub

	Private Sub Class_Terminate
		ReDim arr(-1)
	End Sub

	Public Function First()
		First = arr(0)
	End Function
End Class

Dim p
Set p = New Probe
Response.Write p.First() & "|"
Response.Write TypeName(Eval("true")) & "|"
Response.Write p.First()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok|Boolean|ok" {
		t.Fatalf("unexpected Eval termination side-effect output: %q", output.String())
	}
}

// TestASPCDblConvertsString verifies CDbl parses string numbers used by JSON parser.
func TestASPCDblConvertsString(t *testing.T) {
	source := `<%
Response.Write CDbl("123.456")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "123.456" {
		t.Fatalf("unexpected CDbl output: %q", output.String())
	}
}

// TestASPTypeNameReturnsDate verifies TypeName recognizes date values.
func TestASPTypeNameReturnsDate(t *testing.T) {
	source := `<%
Response.Write TypeName(Now)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Date" {
		t.Fatalf("unexpected TypeName(Now) output: %q", output.String())
	}
}

// TestASPBareRndAutoCallsInValueContext verifies bare zero-arg builtins are
// auto-invoked in arithmetic/value contexts, matching VBScript semantics.
func TestASPBareRndAutoCallsInValueContext(t *testing.T) {
	source := `<%
Randomize 1
Dim code
code = Int((26) * Rnd + 97)
Response.Write code & "|" & Chr(code)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	parts := strings.Split(output.String(), "|")
	if len(parts) != 2 {
		t.Fatalf("unexpected Rnd output format: %q", output.String())
	}
	code, err := strconv.Atoi(parts[0])
	if err != nil {
		t.Fatalf("failed to parse generated code %q: %v", parts[0], err)
	}
	if code < 97 || code > 122 {
		t.Fatalf("expected lowercase ASCII code 97-122, got %d (%q)", code, output.String())
	}
	if parts[1] == "{" {
		t.Fatalf("unexpected random character output %q", output.String())
	}
	if len(parts[1]) != 1 {
		t.Fatalf("expected single generated character, got %q", output.String())
	}
}

// TestASPBareZeroArgGlobalFunctionAutoCallsInValueContext verifies a bare
// zero-argument user Function is auto-invoked when used as a value expression.
func TestASPBareZeroArgGlobalFunctionAutoCallsInValueContext(t *testing.T) {
	source := `<%
Class LoginBox
	Public Function Read()
		Read = cId & "X"
	End Function
End Class

Function cId()
	cId = 1
End Function

Dim l
Set l = New LoginBox
Response.Write l.Read()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "1X" {
		t.Fatalf("expected 1X output, got %q", output.String())
	}
}

// TestASPClassExpressionAutoCallsZeroArgGlobalFunctionBeforeSessionIndex verifies
// the QuickerSite-shaped expression `convertBool(Session(cId & "..."))` works when
// cId is a later zero-argument global Function used bare in value context.
func TestASPClassExpressionAutoCallsZeroArgGlobalFunctionBeforeSessionIndex(t *testing.T) {
	source := `<%
Class cls_LogonEdit
	Public Function logonAdmin()
		logonAdmin = convertBool(Session(cId & "isAUTHENTICATEDasADMIN"))
	End Function
End Class

Function convertBool(value)
	If IsEmpty(value) Then
		convertBool = False
		Exit Function
	End If
	convertBool = CBool(value)
	End Function

Function cId()
	cId = 1
	End Function

Dim logon
Set logon = New cls_LogonEdit
Session("1isAUTHENTICATEDasADMIN") = True
Response.Write CStr(logon.logonAdmin())
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "True" {
		t.Fatalf("expected True output, got %q", output.String())
	}
}

// TestASPBareZeroArgGlobalFunctionAutoCallsInMemberArg verifies that a bare
// zero-arg global Function used as one member-call argument is auto-invoked,
// matching Classic ASP/IIS behavior (e.g., customer.pick(cId)).
func TestASPBareZeroArgGlobalFunctionAutoCallsInMemberArg(t *testing.T) {
	source := `<%
Class Customer
	Public picked
	Public Sub Pick(id)
		picked = id
	End Sub
End Class

Function cId()
	cId = 73
End Function

Dim customer
Set customer = New Customer
customer.Pick(cId)
Response.Write customer.picked
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "73" {
		t.Fatalf("expected member arg output 73, got %q", output.String())
	}
}

// TestASPBareZeroArgGlobalFunctionAutoCallsBeforeMemberAccess verifies that one
// bare zero-arg global Function returning an object is auto-invoked before
// chained member access, matching IIS/VBScript semantics.
func TestASPBareZeroArgGlobalFunctionAutoCallsBeforeMemberAccess(t *testing.T) {
	source := `<%
Class PageObj
	Public iId
End Class

Function getIntranetHomePage()
	Dim page
	Set page = New PageObj
	page.iId = 73
	Set getIntranetHomePage = page
End Function

Function encodeValue(value)
	encodeValue = CStr(value)
End Function

Response.Write encodeValue(getIntranetHomePage.iId)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "73" {
		t.Fatalf("expected member access output 73, got %q", output.String())
	}
}

// TestASPImplicitClassStatementCallByRefObject verifies same-class calls without parentheses support ByRef object out params.
func TestASPImplicitClassStatementCallByRefObject(t *testing.T) {
	source := `<%
Class Widget
End Class

Class Holder
	Private Sub Fill(outObj)
		Set outObj = New Widget
	End Sub

	Public Function HasObject()
		Dim tempObj
		Fill tempObj
		HasObject = IsObject(tempObj) And TypeName(tempObj) = "Widget"
	End Function
End Class

Dim h
Set h = New Holder
Response.Write h.HasObject()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "True" {
		t.Fatalf("unexpected output: %q", output.String())
	}
}

// TestASPRecursiveClassMethodCallWithArgs verifies recursive method calls inside classes bind to the method, not the return variable slot.
func TestASPRecursiveClassMethodCallWithArgs(t *testing.T) {
	source := `<%
Class Node
	Public Function Wrap(level)
		If level <= 0 Then
			Wrap = "x"
		Else
			Wrap = "[" & Wrap(level - 1) & "]"
		End If
	End Function
End Class

Dim n
Set n = New Node
Response.Write n.Wrap(2)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "[[x]]" {
		t.Fatalf("unexpected output: %q", output.String())
	}
}

// TestASPCompileNestedClassDeclarations verifies nested class blocks are rejected.
// Classic ASP/VBScript does not support Class declarations inside another Class body.
func TestASPCompileNestedClassDeclarations(t *testing.T) {
	source := `<html><body><%
Class Outer
	Class Inner
	End Class
End Class
Response.Write "ok"
%></body></html>`

	compiler := NewASPCompiler(source)
	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compile error for nested class declaration, got nil")
	}

	var syntaxErr *vbscript.VBSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected VBSyntaxError, got %T: %v", err, err)
	}
	if syntaxErr.Code != vbscript.SyntaxError {
		t.Fatalf("expected SyntaxError, got code %d (%s)", syntaxErr.Code, syntaxErr.Description)
	}
}

// TestASPClassNewProducesObject verifies New ClassName allocates VTObject instances recognized by IsObject.
func TestASPClassNewProducesObject(t *testing.T) {
	source := `<%
Class Widget
End Class

Dim instance
Set instance = New Widget
If IsObject(instance) Then
	Response.Write "obj"
Else
	Response.Write "no"
End If
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "obj" {
		t.Fatalf("expected obj output, got %q", output.String())
	}
}

// TestASPClassNewAllocatesDistinctInstances verifies each New call receives a distinct instance identity.
func TestASPClassNewAllocatesDistinctInstances(t *testing.T) {
	source := `<%
Class Widget
End Class

Dim a, b
Set a = New Widget
Set b = New Widget

If a Is b Then
	Response.Write "same"
Else
	Response.Write "diff"
End If
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "diff" {
		t.Fatalf("expected diff output, got %q", output.String())
	}
}

// TestASPClassNewUndefinedRaisesRuntimeError verifies undefined class names fail with VBScript class-type runtime error.
func TestASPClassNewUndefinedRaisesRuntimeError(t *testing.T) {
	source := `<%
Dim instance
Set instance = New MissingClass
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	err := vm.Run()
	if err == nil {
		t.Fatalf("expected runtime error for undefined class")
	}

	var runtimeErr *VMError
	if !errors.As(err, &runtimeErr) {
		t.Fatalf("expected VMError, got %T", err)
	}
	if runtimeErr.Code != vbscript.ClassTypeIsNotDefined {
		t.Fatalf("unexpected runtime code: got %d want %d", runtimeErr.Code, vbscript.ClassTypeIsNotDefined)
	}
}

// TestASPClassMethodFunctionCall verifies class Function dispatch through obj.Method(args) call syntax.
func TestASPClassMethodFunctionCall(t *testing.T) {
	source := `<%
Class MathBox
	Public Function Add(a, b)
		Add = a + b
	End Function
End Class

Dim m
Set m = New MathBox
Response.Write m.Add(2, 3)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "5" {
		t.Fatalf("expected 5 output, got %q", output.String())
	}
}

// TestASPClassMethodCallCaseInsensitive verifies class method dispatch ignores method-name casing.
func TestASPClassMethodCallCaseInsensitive(t *testing.T) {
	source := `<%
Class TextBox
	Public Function Echo(value)
		Echo = value
	End Function
End Class

Dim t
Set t = New TextBox
Response.Write t.eChO("Axon")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Axon" {
		t.Fatalf("expected Axon output, got %q", output.String())
	}
}

// TestASPClassMethodSubStatementCall verifies statement-style class Sub calls without parentheses.
func TestASPClassMethodSubStatementCall(t *testing.T) {
	source := `<%
Class Greeter
	Public Sub WriteHello(target)
		Response.Write "Hi " & target
	End Sub
End Class

Dim g
Set g = New Greeter
g.WriteHello "VM"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Hi VM" {
		t.Fatalf("expected Hi VM output, got %q", output.String())
	}
}

// TestASPClassMethodMissingRaisesRuntimeError verifies unknown class methods raise VBScript object method errors.
func TestASPClassMethodMissingRaisesRuntimeError(t *testing.T) {
	source := `<%
Class Greeter
	Public Sub Ping()
	End Sub
End Class

Dim g
Set g = New Greeter
g.UnknownMethod
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	err := vm.Run()
	if err == nil {
		t.Fatalf("expected runtime error for missing class method")
	}

	var runtimeErr *VMError
	if !errors.As(err, &runtimeErr) {
		t.Fatalf("expected VMError, got %T", err)
	}
	if runtimeErr.Code != vbscript.ObjectDoesntSupportThisPropertyOrMethod {
		t.Fatalf("unexpected runtime code: got %d want %d", runtimeErr.Code, vbscript.ObjectDoesntSupportThisPropertyOrMethod)
	}
}

// TestASPClassPropertyDeclarationsCompile verifies Property Get/Let/Set blocks compile inside classes.
func TestASPClassPropertyDeclarationsCompile(t *testing.T) {
	source := `<%
Class Counter
	Private m_value

	Public Property Get Value()
		Value = m_value
	End Property

	Public Property Let Value(v)
		m_value = v
	End Property

	Public Property Set RefValue(obj)
		Set m_value = obj
	End Property
End Class

Response.Write "ok"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected ok output, got %q", output.String())
	}
}

// TestASPClassPropertySignatureMismatchFailsCompile verifies Property Get/Let signature mismatch raises compile error.
func TestASPClassPropertySignatureMismatchFailsCompile(t *testing.T) {
	source := `<%
Class Counter
	Public Property Get Value(index)
		Value = index
	End Property

	Public Property Let Value(v)
	End Property
End Class
%>`

	compiler := NewASPCompiler(source)
	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compile error for property signature mismatch")
	}

	var syntaxErr *vbscript.VBSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected VBSyntaxError, got %T", err)
	}
}

// TestASPClassPropertyLetWithoutValueParameterFailsCompile verifies Property Let requires one value parameter.
func TestASPClassPropertyLetWithoutValueParameterFailsCompile(t *testing.T) {
	source := `<%
Class Counter
	Public Property Let Value()
	End Property
End Class
%>`

	compiler := NewASPCompiler(source)
	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compile error for Property Let without value parameter")
	}

	var syntaxErr *vbscript.VBSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected VBSyntaxError, got %T", err)
	}
}

// TestASPClassPropertyGetLetRuntime verifies property Let assignment and property Get read dispatch.
func TestASPClassPropertyGetLetRuntime(t *testing.T) {
	source := `<%
Class Counter
	Private m_value

	Public Property Get Value()
		Value = m_value
	End Property

	Public Property Let Value(v)
		m_value = v
	End Property
End Class

Dim c
Set c = New Counter
c.Value = 7
Response.Write c.Value
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "7" {
		t.Fatalf("expected 7 output, got %q", output.String())
	}
}

// TestASPDefaultPropertyPageSemanticsAlignment locks in the ASP behavior used by
// www/tests/test_default_property.asp so page-level tests stay aligned with VM semantics.
func TestASPDefaultPropertyPageSemanticsAlignment(t *testing.T) {
	source := `<%
Class SimpleCounter
	Private Count

	Public Default Property Get Value()
		Value = Count
	End Property

	Public Default Property Let Value(newVal)
		Count = newVal
	End Property

	Public Sub Increment()
		Count = Count + 1
	End Sub
End Class

Class StringStore
	Private Items()
	Private itemCount

	Public Sub Initialize()
		ReDim Items(9)
		itemCount = 0
	End Sub

	Public Default Property Get Item(index)
		If index >= 0 And index < itemCount Then
			Item = Items(index)
		Else
			Item = Empty
		End If
	End Property

	Public Default Property Let Item(index, val)
		If index >= 0 Then
			If index >= itemCount Then
				itemCount = index + 1
			End If
			Items(index) = val
		End If
	End Property

	Public Function Count()
		Count = itemCount
	End Function
End Class

Dim counter
Set counter = New SimpleCounter
Response.Write counter & "|"
counter.Value = 42
Response.Write counter & "|"
counter.Increment()
Response.Write counter & "|"
counter.Value = counter + 8
Response.Write counter & "|"

Dim store
Set store = New StringStore
store.Initialize()
store(0) = "First"
store(1) = "Second"
store(2) = "Third"
Response.Write store(0) & "|" & store(1) & "|" & store(2) & "|" & store.Count()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "|42|43|51|First|Second|Third|3"
	if output.String() != expected {
		t.Fatalf("unexpected default-property alignment output: got %q want %q", output.String(), expected)
	}
}

// TestASPClassPropertySetObjectRuntime verifies explicit Set assignment routes to Property Set.
func TestASPClassPropertySetObjectRuntime(t *testing.T) {
	source := `<%
Class Widget
End Class

Class Holder
	Private m_item

	Public Property Set Item(v)
		Set m_item = v
	End Property

	Public Property Get Item()
		Set Item = m_item
	End Property
End Class

Dim h, w
Set h = New Holder
Set w = New Widget
Set h.Item = w

If h.Item Is w Then
	Response.Write "ok"
Else
	Response.Write "no"
End If
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected ok output, got %q", output.String())
	}
}

// TestASPClassPropertySetExplicitWithoutAccessorFails verifies explicit Set requires Property Set accessor.
func TestASPClassPropertySetExplicitWithoutAccessorFails(t *testing.T) {
	source := `<%
Class Widget
End Class

Class Holder
	Private m_value

	Public Property Let Item(v)
		m_value = v
	End Property
End Class

Dim h, w
Set h = New Holder
Set w = New Widget
Set h.Item = w
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	err := vm.Run()
	if err == nil {
		t.Fatalf("expected runtime error for explicit Set without Property Set accessor")
	}

	var runtimeErr *VMError
	if !errors.As(err, &runtimeErr) {
		t.Fatalf("expected VMError, got %T", err)
	}
	if runtimeErr.Code != vbscript.WrongNumberOfParameters {
		t.Fatalf("unexpected runtime code: got %d want %d", runtimeErr.Code, vbscript.WrongNumberOfParameters)
	}
}

// TestASPClassPrivateMembersAreHidden verifies private class members are not callable from outside.
func TestASPClassPrivateMembersAreHidden(t *testing.T) {
	source := `<%
Class Secret
	Private Function Hidden()
		Hidden = 42
	End Function

	Private Property Get Code()
		Code = 7
	End Property
End Class

Dim s
Set s = New Secret
Response.Write s.Hidden()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	err := vm.Run()
	if err == nil {
		t.Fatalf("expected runtime error for private member access")
	}

	var runtimeErr *VMError
	if !errors.As(err, &runtimeErr) {
		t.Fatalf("expected VMError, got %T", err)
	}
	if runtimeErr.Code != vbscript.ObjectDoesntSupportThisPropertyOrMethod {
		t.Fatalf("unexpected runtime code: got %d want %d", runtimeErr.Code, vbscript.ObjectDoesntSupportThisPropertyOrMethod)
	}
}

// TestASPClassInitializeRunsOnNew verifies Class_Initialize is invoked during object construction.
func TestASPClassInitializeRunsOnNew(t *testing.T) {
	source := `<%
Class Widget
	Private Sub Class_Initialize()
		Response.Write "I"
	End Sub
End Class

Dim w
Set w = New Widget
Response.Write "K"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "IK" {
		t.Fatalf("expected IK output, got %q", output.String())
	}
}

// TestASPClassTerminateRunsAfterScript verifies Class_Terminate is invoked at deterministic VM cleanup.
func TestASPClassTerminateRunsAfterScript(t *testing.T) {
	source := `<%
Class Widget
	Private Sub Class_Terminate()
		Response.Write "T"
	End Sub
End Class

Dim w
Set w = New Widget
Response.Write "K"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "KT" {
		t.Fatalf("expected KT output, got %q", output.String())
	}
}

// TestASPClassInitializeWithArgumentsFailsCompile verifies lifecycle methods cannot declare parameters.
func TestASPClassInitializeWithArgumentsFailsCompile(t *testing.T) {
	source := `<%
Class Widget
	Private Sub Class_Initialize(v)
	End Sub
End Class
%>`

	compiler := NewASPCompiler(source)
	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compile error for Class_Initialize with arguments")
	}

	var syntaxErr *vbscript.VBSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected VBSyntaxError, got %T", err)
	}
	if syntaxErr.Code != vbscript.ClassInitializeOrTerminateDoNotHaveArguments {
		t.Fatalf("unexpected compile code: got %d want %d", syntaxErr.Code, vbscript.ClassInitializeOrTerminateDoNotHaveArguments)
	}
}

// TestASPClassMethodWrongArgumentCountFails verifies class method calls enforce parameter count with VBScript code 450.
func TestASPClassMethodWrongArgumentCountFails(t *testing.T) {
	source := `<%
Class MathBox
	Public Function Add(a, b)
		Add = a + b
	End Function
End Class

Dim m
Set m = New MathBox
Response.Write m.Add(1)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	err := vm.Run()
	if err == nil {
		t.Fatalf("expected runtime error for wrong argument count")
	}

	var runtimeErr *VMError
	if !errors.As(err, &runtimeErr) {
		t.Fatalf("expected VMError, got %T", err)
	}
	if runtimeErr.Code != vbscript.WrongNumberOfParameters {
		t.Fatalf("unexpected runtime code: got %d want %d", runtimeErr.Code, vbscript.WrongNumberOfParameters)
	}
}

// TestASPClassArrayRedimInMethod verifies Dim/ReDim operations inside class methods follow VBScript array behavior.
func TestASPClassArrayRedimInMethod(t *testing.T) {
	source := `<%
Class Bag
	Public Function Build()
		Dim values()
		ReDim values(0)
		values(0) = "A"
		ReDim Preserve values(1)
		values(1) = "B"
		Build = values(0) & values(1)
	End Function
End Class

Dim b
Set b = New Bag
Response.Write b.Build()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "AB" {
		t.Fatalf("expected AB output, got %q", output.String())
	}
}

// TestASPClassInitializeBuildsInstanceArray verifies Class_Initialize can populate per-instance array state.
func TestASPClassInitializeBuildsInstanceArray(t *testing.T) {
	source := `<%
Class TestArray2
	Dim i_items, i_count, i_capacity

	Private Sub Class_Initialize()
		ReDim i_items(-1)
		i_count = 0
		i_capacity = 0
	End Sub

	Public Property Get Items()
		Dim tmp
		tmp = i_items
		If i_count < i_capacity Then
			ReDim Preserve tmp(i_count - 1)
		End If
		Items = tmp
	End Property
End Class

Dim obj
Set obj = New TestArray2
Dim result
result = obj.Items
Response.Write TypeName(result)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Variant()" {
		t.Fatalf("expected Variant() output, got %q", output.String())
	}
}

// TestASPClassDirectFieldAccess verifies direct class fields declared in the class body can be read and written externally.
func TestASPClassDirectFieldAccess(t *testing.T) {
	source := `<%
Class SimpleClass
	Dim myname
End Class

Dim obj
Set obj = New SimpleClass
obj.myname = "test"
Response.Write obj.myname
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "test" {
		t.Fatalf("expected test output, got %q", output.String())
	}
}

// TestASPClassDimFieldPersistsAcrossSetValueGetValue verifies class-level Dim fields
// persist correctly when written/read through class methods.
func TestASPClassDimFieldPersistsAcrossSetValueGetValue(t *testing.T) {
	source := `<%
Class Counter
	Dim value

	Public Sub SetValue(v)
		value = v
	End Sub

	Public Function GetValue()
		GetValue = value
	End Function
End Class

Dim c
Set c = New Counter
c.SetValue 42
Response.Write c.GetValue()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "42" {
		t.Fatalf("expected 42 output, got %q", output.String())
	}
}

// TestASPClassSetIntoMemberArray verifies Set member(index) = obj works for per-instance array members.
func TestASPClassSetIntoMemberArray(t *testing.T) {
	source := `<%
Class JSONpair
	Dim name, value
End Class

Class TestAddComplex
	Dim i_properties, i_count, i_capacity

	Private Sub Class_Initialize()
		ReDim i_properties(-1)
		i_count = 0
		i_capacity = 0
	End Sub

	Public Sub Add(ByVal prop, ByVal obj)
		Dim item
		Set item = New JSONpair
		item.name = prop
		item.value = obj

		If i_count >= i_capacity Then
			ReDim Preserve i_properties(i_capacity * 1.2 + 1)
			i_capacity = UBound(i_properties) + 1
		End If

		Set i_properties(i_count) = item
		i_count = i_count + 1
	End Sub

	Public Function FirstName()
		FirstName = i_properties(0).name
	End Function
End Class

Dim obj
Set obj = New TestAddComplex
obj.Add "test", "value"
Response.Write obj.FirstName()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "test" {
		t.Fatalf("expected test output, got %q", output.String())
	}
}

// TestASPClassCompareWithInitialize verifies class field comparisons after Class_Initialize match IIS-compatible behavior.
func TestASPClassCompareWithInitialize(t *testing.T) {
	source := `<%
Class TestCompare
	Dim i_count, i_capacity

	Private Sub Class_Initialize()
		i_count = 5
		i_capacity = 10
	End Sub

	Public Sub TestMethod()
		If i_count >= i_capacity Then
			Response.Write "Count is less than capacity<br>"
		Else
			Response.Write "Count is greater than or equal to capacity<br>"
		End If
	End Sub
End Class

Dim obj
Set obj = New TestCompare
obj.TestMethod()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Count is greater than or equal to capacity<br>" {
		t.Fatalf("unexpected output: %q", output.String())
	}
}

// TestASPOnErrorMissingCreateObjectRecordsetEOF verifies missing COM object creation does not trigger runaway rs.EOF evaluation loops.
func TestASPOnErrorMissingCreateObjectRecordsetEOF(t *testing.T) {
	source := `<%
On Error Resume Next
Dim conn, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = conn.Execute("SELECT 1")

If Not rs.EOF Then
	Response.Write "loop"
End If
Response.Write "done"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "done" {
		t.Fatalf("expected done output, got %q", output.String())
	}
}

// TestASPSetClassMemberFromForwardFunction ensures "Set field = funcName" inside
// class methods resolves funcName as a class member call even when declared later.
func TestASPSetClassMemberFromForwardFunction(t *testing.T) {
	source := `<%
Option Explicit
Class C
	Private plugins
	Private Sub Class_Initialize()
		Set plugins = Nothing
	End Sub

	Public Sub A()
		If plugins Is Nothing Then Set plugins = dict
		Response.Write "T=" & TypeName(plugins) & "|E=" & Err.Number
	End Sub

	Public Function dict()
		Set dict = Server.CreateObject("Scripting.Dictionary")
	End Function
End Class

Dim c
Set c = New C
c.A
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "T=Dictionary|E=0" {
		t.Fatalf("unexpected output: %q", output.String())
	}
}

// TestASPStatementResponseWrite verifies statement-style member call emits output.
func TestASPStatementResponseWrite(t *testing.T) {
	source := `<%
Response.Write "ok"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected statement Response.Write output, got %q", output.String())
	}
}

// TestASPCompileSupportsSetCreateObject verifies Set assignments with Server.CreateObject compile correctly.
func TestASPCompileSupportsSetCreateObject(t *testing.T) {
	source := `<%
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
}

// TestASPCompileSupportsObjectPropertySet verifies obj.Property = value compiles correctly (OpMemberSet).
func TestASPCompileSupportsObjectPropertySet(t *testing.T) {
	source := `<%
Dim g3md
Set g3md = Server.CreateObject("G3Md")
g3md.Unsafe = True
g3md.HardWraps = False
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
}

// TestASPFunctionCallBeforeDeclaration verifies forward calls to user functions declared later in ASP code.
func TestASPFunctionCallBeforeDeclaration(t *testing.T) {
	source := `<%
Dim page, mapped
page = Request.QueryString("page")
mapped = Server.MapPath("menu.md")
Response.Write ReadFile(page, mapped)

Function ReadFile(value, physicalPath)
	ReadFile = value & "|" & physicalPath
End Function
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir("./www")
	host.Server().SetRequestPath("/manual/default.asp")
	host.Request().QueryString.Add("page", "intro/welcome")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	rendered := output.String()
	if !strings.Contains(rendered, "intro/welcome|") {
		t.Fatalf("expected QueryString value in output, got %q", rendered)
	}
	if !strings.Contains(strings.ReplaceAll(rendered, "\\", "/"), "/www/manual/menu.md") {
		t.Fatalf("expected MapPath result in output, got %q", rendered)
	}
}

// TestASPManualPageQueryAndReplace verifies Request.QueryString and Replace flow used by manual/default.asp.
func TestASPManualPageQueryAndReplace(t *testing.T) {
	source := `<%
Dim page
page = Request.QueryString("page")
If page = "" Then page = "intro/welcome"
page = Replace(page, "..", "")
Response.Write page
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Request().QueryString.Add("page", "docs/getting-started")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "docs/getting-started" {
		t.Fatalf("expected query string page value, got %q", output.String())
	}
}

// TestASPManualPageDefaultFallback verifies page fallback when query string is empty.
func TestASPManualPageDefaultFallback(t *testing.T) {
	source := `<%
Dim page
page = Request.QueryString("page")
If page = "" Then page = "intro/welcome"
page = Replace(page, "..", "")
Response.Write page
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "intro/welcome" {
		t.Fatalf("expected default page fallback, got %q", output.String())
	}
}

// TestASPQueryStringExpression verifies Request.QueryString("page") in expression context.
func TestASPQueryStringExpression(t *testing.T) {
	source := `<%= Request.QueryString("page") %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Request().QueryString.Add("page", "alpha")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if output.String() != "alpha" {
		t.Fatalf("expected alpha from Request.QueryString, got %q", output.String())
	}
}

// TestASPReplaceBuiltin verifies Replace builtin registration and execution.
func TestASPReplaceBuiltin(t *testing.T) {
	source := `<%= Replace("abc..def", "..", "") %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if output.String() != "abcdef" {
		t.Fatalf("expected Replace result abcdef, got %q", output.String())
	}
}

// TestASPSelectCaseBasic verifies Select Case with direct value matching and Case Else.
func TestASPSelectCaseBasic(t *testing.T) {
	source := `<%
Dim value
value = 2
Select Case value
Case 1
	Response.Write "one"
Case 2, 3
	Response.Write "two-or-three"
Case Else
	Response.Write "other"
End Select
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "two-or-three" {
		t.Fatalf("expected Select Case output two-or-three, got %q", output.String())
	}
}

// TestASPSelectCaseRangeAndElse verifies Case x To y range matching and Case Else fallback.
func TestASPSelectCaseRangeAndElse(t *testing.T) {
	source := `<%
Dim grade
grade = 85

Select Case grade
Case 90 To 100
	Response.Write "A"
Case 80 To 89
	Response.Write "B"
Case Else
	Response.Write "F"
End Select
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "B" {
		t.Fatalf("expected Select Case range output B, got %q", output.String())
	}
}

// TestASPSelectCaseNoFallthrough verifies VBScript Select Case exits after first matching Case.
func TestASPSelectCaseNoFallthrough(t *testing.T) {
	source := `<%
Dim x
x = 2

Select Case x
Case 2
	Response.Write "first"
Case 2, 3
	Response.Write "second"
Case Else
	Response.Write "else"
End Select
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "first" {
		t.Fatalf("expected no-fallthrough output first, got %q", output.String())
	}
}

// TestASPSelectCaseVariantCoercion verifies Select Case comparisons follow VBScript Variant coercion rules.
func TestASPSelectCaseVariantCoercion(t *testing.T) {
	source := `<%
Dim a, b

a = "1"
Select Case a
Case 1
	Response.Write "num-match|"
Case Else
	Response.Write "num-miss|"
End Select

b = True
Select Case b
Case False
	Response.Write "bool-miss"
Case True
	Response.Write "bool-match"
Case Else
	Response.Write "bool-else"
End Select
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "num-match|bool-match" {
		t.Fatalf("expected Variant-coercion output num-match|bool-match, got %q", output.String())
	}
}

// TestASPOnErrorResumeNextKeywordNext verifies On Error Resume Next recognizes keyword token Next.
func TestASPOnErrorResumeNextKeywordNext(t *testing.T) {
	source := `<%
On Error Resume Next
Dim x
x = 1 / 0
Response.Write "after"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "after" {
		t.Fatalf("expected On Error Resume Next output after, got %q", output.String())
	}
}

// TestASPBracketEscapedBuiltinIdentifierCompilesAndRuns verifies that bracket-escaped
// identifiers ([isEmpty], [IsNull]) resolve to the native VBScript builtins.
// VBScript IsEmpty returns True only for the uninitialized Empty variant, not for "".
// VBScript IsNull returns True only for the Null variant.
func TestASPBracketEscapedBuiltinIdentifierCompilesAndRuns(t *testing.T) {
	source := `<%
Response.Write [isEmpty](Empty)
Response.Write "|"
Response.Write [IsNull](Null)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "True|True" {
		t.Fatalf("expected bracket-escaped builtin output True|True, got %q", output.String())
	}
}

// TestASPBracketEscapedIsEmptyEmptyStringReturnsFalse verifies that [isEmpty]("") resolves
// to the native VBScript IsEmpty builtin, which returns False for an empty string (not Empty).
func TestASPBracketEscapedIsEmptyEmptyStringReturnsFalse(t *testing.T) {
	source := `<%
Response.Write [isEmpty]("")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	// IsEmpty("") is False: empty string is not the VBScript Empty (uninitialized) variant.
	if output.String() != "False" {
		t.Fatalf("expected [isEmpty](\"\") == False, got %q", output.String())
	}
}

// TestASPBracketEscapedIsNullNonNullReturnsFalse verifies that [IsNull](0) returns False.
func TestASPBracketEscapedIsNullNonNullReturnsFalse(t *testing.T) {
	source := `<%
Response.Write [IsNull](0)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "False" {
		t.Fatalf("expected [IsNull](0) == False, got %q", output.String())
	}
}

// TestASPOnErrorGoto0CompileAndRun verifies that "On Error GoTo 0" compiles successfully
// and clears the error handler (any subsequent runtime error propagates immediately).
func TestASPOnErrorGoto0CompileAndRun(t *testing.T) {
	source := `<%
On Error Resume Next
Err.Raise 5
On Error GoTo 0
Response.Write "ok"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("On Error GoTo 0 should compile without error, got: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected output ok, got %q", output.String())
	}
}

// TestASPOnErrorGotoLabelIsSyntaxError verifies that "On Error GoTo <label>" (any non-zero
// label) is rejected at compile time with a VBScript SyntaxError, because VBScript only
// supports "On Error GoTo 0" and "On Error Resume Next".
func TestASPOnErrorGotoLabelIsSyntaxError(t *testing.T) {
	cases := []string{
		"On Error GoTo MyHandler",
		"On Error GoTo errHandler",
		"On Error GoTo 1",
	}
	for _, src := range cases {
		compiler := NewASPCompiler("<%" + src + "%>")
		err := compiler.Compile()
		if err == nil {
			t.Fatalf("[%s] expected compile error, got none", src)
		}
		var syntaxErr *vbscript.VBSyntaxError
		if !errors.As(err, &syntaxErr) {
			t.Fatalf("[%s] expected VBSyntaxError, got %T: %v", src, err, err)
		}
		if syntaxErr.Code != vbscript.SyntaxError {
			t.Fatalf("[%s] expected SyntaxError code %d, got %d", src, vbscript.SyntaxError, syntaxErr.Code)
		}
	}
}

// TestASPObjectTokenRegistersApplicationStaticObject verifies <object runat="server"> populates Application.StaticObjects.
func TestASPObjectTokenRegistersApplicationStaticObject(t *testing.T) {
	source := `<object runat="server" scope="Application" id="demoObj" progid="Scripting.FileSystemObject"></object>
<%= Application.StaticObjects.Count %>|<%= IsObject(Application.StaticObjects("demoObj")) %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "\n1|True" {
		t.Fatalf("expected Application.StaticObjects registration output \\n1|True, got %q", output.String())
	}
}

// TestASPObjectTokenRegistersSessionStaticObject verifies Session.StaticObjects member chaining.
func TestASPObjectTokenRegistersSessionStaticObject(t *testing.T) {
	source := `<object runat="server" scope="Session" id="sessObj" progid="Scripting.Dictionary"></object>
<%= Session.StaticObjects.Count %>|<%= IsObject(Session.StaticObjects.Item("sessObj")) %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "\n1|True" {
		t.Fatalf("expected Session.StaticObjects registration output \\n1|True, got %q", output.String())
	}
}

// TestASPObjectTokenMaterializesSimulatedActiveX verifies that global static-object
// markers instantiate simulated libraries beyond Scripting.Dictionary.
func TestASPObjectTokenMaterializesSimulatedActiveX(t *testing.T) {
	source := `<object runat="server" scope="Application" id="jsonObj" progid="G3JSON"></object>
<%= IsObject(Application.StaticObjects("jsonObj")) %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "\nTrue" {
		t.Fatalf("expected simulated ActiveX static object output \\nTrue, got %q", output.String())
	}
}

// TestASPImplicitStaticObjectIdentifierFallback verifies undeclared identifier reads in ASP
// resolve static objects by ID from Session.StaticObjects first, then Application.StaticObjects.
func TestASPImplicitStaticObjectIdentifierFallback(t *testing.T) {
	source := `<%
Response.Write CStr(IsObject(appObj)) & "|"
Response.Write CStr(IsObject(sessObj)) & "|"
appObj.Add "x", "1"
sessObj.Add "y", "1"
Response.Write CStr(appObj.Count) & "|"
Response.Write CStr(sessObj.Count)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	app := host.Application()
	session := host.Session()
	app.AddStaticObject("appObj", asp.NewApplicationString(staticObjectProgIDPrefix+"Scripting.Dictionary"))
	session.AddStaticObject("sessObj", asp.NewApplicationString(staticObjectProgIDPrefix+"Scripting.Dictionary"))

	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "True|True|1|1" {
		t.Fatalf("expected static object fallback output True|True|1|1, got %q", output.String())
	}
}

// TestASPPreprocessStripsEmptyASPBlocks verifies requested cleanup for empty ASP tags.
func TestASPPreprocessStripsEmptyASPBlocks(t *testing.T) {
	source := `A<%=%>B<%= %>C<%%>D<% %>E`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ABCDE" {
		t.Fatalf("expected output ABCDE after stripping empty ASP blocks, got %q", output.String())
	}
}

// TestASPCommentIgnoresASPMarkersInLine verifies markers like <% and <%= inside
// VBScript apostrophe comments are treated as plain comment text. The closing %>
// delimiter itself is not swallowed, matching Classic ASP behaviour.
func TestASPCommentIgnoresASPMarkersInLine(t *testing.T) {
	source := `<%
' Comment with markers: <%=var text <% more text
Response.Write "ok"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("expected comment-marker output ok, got %q", output.String())
	}
}

// TestASPCommentIgnoresPercentCodeEndMarker verifies %> inside apostrophe comments
// does not terminate the current ASP code block.
func TestASPCommentIgnoresPercentCodeEndMarker(t *testing.T) {
	source := `<%
Dim test1
test1 = "before"
' This is a comment with %> in the middle and more text after
test1 = test1 & " after"
Response.Write test1
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "before after" {
		t.Fatalf("expected output before after, got %q", output.String())
	}
}

// TestASPCommentIgnoresMultiplePercentCodeEndMarkers verifies multiple %> sequences
// inside a comment remain comment text until end of line.
func TestASPCommentIgnoresMultiplePercentCodeEndMarkers(t *testing.T) {
	source := `<%
Dim test2
test2 = "start"
' Comment with multiple markers: %> text <%=var%> more text %>
test2 = test2 & " end"
Response.Write test2
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "start end" {
		t.Fatalf("expected output start end, got %q", output.String())
	}
}

// TestLexerASPObjectTokenAttributes verifies object declaration attributes are extracted correctly.
func TestLexerASPObjectTokenAttributes(t *testing.T) {
	lex := vbscript.NewLexer(`<object runat="server" scope="Application" id="demoObj" progid="Scripting.FileSystemObject"></object>`)
	lex.Mode = vbscript.ModeASP
	for {
		token := lex.NextToken()
		if objectToken, ok := token.(*vbscript.ASPObjectToken); ok {
			if !strings.EqualFold(objectToken.Scope, "Application") {
				t.Fatalf("unexpected object scope: %q", objectToken.Scope)
			}
			if objectToken.ID != "demoObj" {
				t.Fatalf("unexpected object id: %q", objectToken.ID)
			}
			if !strings.EqualFold(objectToken.ProgID, "Scripting.FileSystemObject") {
				t.Fatalf("unexpected object progid: %q", objectToken.ProgID)
			}
			return
		}
		if _, ok := token.(*vbscript.EOFToken); ok {
			break
		}
	}

	t.Fatalf("expected ASPObjectToken, got EOF")
}

// TestASPDoLoopVariants verifies classic Do/Loop forms (pre/post While/Until) compile and execute.
func TestASPDoLoopVariants(t *testing.T) {
	source := `<%
Dim i, out
out = ""

' Do Until (pre-test)
i = 0
Do Until i = 3
    out = out & "A"
    i = i + 1
Loop

' Do While (pre-test)
i = 0
Do While i < 2
    out = out & "B"
    i = i + 1
Loop

' Do ... Loop Until (post-test)
i = 0
Do
    out = out & "C"
    i = i + 1
Loop Until i = 2

' Do ... Loop While (post-test)
i = 0
Do
    out = out & "D"
    i = i + 1
Loop While i < 3

Response.Write out
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "AAABBCCDDD" {
		t.Fatalf("unexpected Do/Loop output: got %q want %q", output.String(), "AAABBCCDDD")
	}
}

// TestASPEmptyEqualsEmptyString verifies VBScript-style Empty to "" comparison used by manual/default.asp.
func TestASPEmptyEqualsEmptyString(t *testing.T) {
	source := `<%
Dim page
If page = "" Then
	Response.Write "T"
Else
	Response.Write "F"
End If
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "T" {
		t.Fatalf("expected Empty = \"\" to be true, got %q", output.String())
	}
}

// TestASPLateDimDoesNotResetAssignedValue verifies VBScript declaration behavior
// where a late scalar Dim declaration does not reset a value assigned earlier.
func TestASPLateDimDoesNotResetAssignedValue(t *testing.T) {
	source := `<%
lateDimProbe = "kept"
Dim lateDimProbe
Response.Write lateDimProbe
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "kept" {
		t.Fatalf("expected late Dim to preserve assignment, got %q", output.String())
	}
}

// TestASPLateLocalDimUsesWholeProcedureScope verifies one late local Dim binds
// earlier reads and writes to the same procedure-local slot, matching IIS/VBScript.
func TestASPLateLocalDimUsesWholeProcedureScope(t *testing.T) {
	source := `<%
Function BuildValue()
	value = "kept"
	Dim value
	BuildValue = value
End Function

Response.Write BuildValue()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "kept" {
		t.Fatalf("expected kept output, got %q", output.String())
	}
}

// TestASPLateLocalDimBeforeSetMemberCallRHS verifies late Dim hoisting preserves
// Set x = obj.Method(localVar) patterns inside class methods.
func TestASPLateLocalDimBeforeSetMemberCallRHS(t *testing.T) {
	source := `<%
Class DBObj
	Public Function Execute(sql)
		Set Execute = Server.CreateObject("Scripting.Dictionary")
	End Function
End Class

Class Holder
	Public Function Read()
		sql = "select 1"
		Dim rs, sql
		Set rs = db.Execute(sql)
		Read = TypeName(rs)
	End Function
End Class

Dim db, h
Set db = New DBObj
Set h = New Holder
Response.Write h.Read()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Dictionary" {
		t.Fatalf("expected Dictionary output, got %q", output.String())
	}
}

// TestASPPropertyGetSeesForwardGlobalObject verifies one Property Get does not
// hoist later page-level Dim declarations into property-local scope.
func TestASPPropertyGetSeesForwardGlobalObject(t *testing.T) {
	source := `<%
Class CustomerObj
	Public iId
End Class

Class Widget
	Public Property Get ShowValueProp
		ShowValueProp = customer.iId
	End Property
End Class

Dim widget
Dim customer
Set customer = New CustomerObj
customer.iId = 73
Set widget = New Widget

Response.Write widget.ShowValueProp
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "73" {
		t.Fatalf("expected 73 output, got %q", output.String())
	}
}

// TestASPForwardZeroArgFunctionArgumentAutoInvokes verifies one bare global
// zero-arg Function used as a call argument before its declaration still passes
// the function result rather than the underlying UserSub slot.
func TestASPForwardZeroArgFunctionArgumentAutoInvokes(t *testing.T) {
	source := `<%
Dim pageBody
pageBody = Replace("x", "x", showSiteMap, 1, -1, 1)
Response.Write pageBody

Function showSiteMap
	showSiteMap = "OK"
End Function
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "OK" {
		t.Fatalf("expected OK output, got %q", output.String())
	}
}

// TestASPForLoopInterleavedHTML verifies HTML/ASP interleaving inside For...Next emits markup on each iteration.
func TestASPForLoopInterleavedHTML(t *testing.T) {
	source := `<%
Dim i
For i = 1 To 3
%><li><%= i %></li><%
Next
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "<li>1</li><li>2</li><li>3</li>" {
		t.Fatalf("expected loop markup output, got %q", output.String())
	}
}

// TestASPForLoopSidebarArrayValues verifies array fill/read in For loops with interleaved ASP/HTML output.
func TestASPForLoopSidebarArrayValues(t *testing.T) {
	source := `<%
Dim testFiles(), testCount, i
testCount = 0
ReDim testFiles(0)

For i = 1 To 3
	If testCount > 0 Then
		ReDim Preserve testFiles(testCount)
	End If
	testFiles(testCount) = "test_" & i & ".asp"
	testCount = testCount + 1
Next

For i = 0 To testCount - 1
%><a href="?file=<%= testFiles(i) %>"><%= testFiles(i) %></a><%
Next
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := `<a href="?file=test_1.asp">test_1.asp</a><a href="?file=test_2.asp">test_2.asp</a><a href="?file=test_3.asp">test_3.asp</a>`
	if output.String() != expected {
		t.Fatalf("expected sidebar-style link output, got %q", output.String())
	}
}

// TestASPClassMeKeyword verifies that the VBScript "Me" keyword correctly refers to the current
// class instance inside methods and properties, allowing access to public members via Me.MemberName.
func TestASPClassMeKeyword(t *testing.T) {
	source := `<%
Class Person
    Public Name

    Private Sub Class_Initialize()
        Name = "Unnamed"
    End Sub

    Public Function SelfRef()
        SelfRef = Me.Name
    End Function

    Public Function FullLabel()
        FullLabel = "Name=" & Me.Name
    End Function
End Class

Dim p
Set p = New Person
p.Name = "Alice"

Response.Write p.SelfRef() & "|"
Response.Write p.FullLabel() & "|"

Dim p2
Set p2 = New Person
p2.Name = "Bob"
Response.Write p2.SelfRef()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	got := output.String()
	if got != "Alice|Name=Alice|Bob" {
		t.Fatalf("expected 'Alice|Name=Alice|Bob', got %q", got)
	}
}

// TestASPScriptingDictionary verifies that Scripting.Dictionary can be created via
// Server.CreateObject and that Add, Exists, Item, Remove, Count, Keys, Items all
// behave identically to classic Microsoft Scripting.Dictionary.
func TestASPScriptingDictionary(t *testing.T) {
	source := `<%
Dim d
Set d = Server.CreateObject("Scripting.Dictionary")

d.Add "Name", "Alice"
d.Add "Age", 30

Response.Write d("Name") & "|"
Response.Write d.Item("Age") & "|"
Response.Write d.Count & "|"
Response.Write d.Exists("Name") & "|"
Response.Write d.Exists("Missing") & "|"

d.Remove "Age"
Response.Write d.Count & "|"

Dim k
k = d.Keys()
Response.Write k(0) & "|"

d.RemoveAll
Response.Write d.Count
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	got := output.String()
	if got != "Alice|30|2|True|False|1|Name|0" {
		t.Fatalf("expected 'Alice|30|2|True|False|1|Name|0', got %q", got)
	}
}

// TestASPScriptingDictionaryClassScope verifies that a Scripting.Dictionary created in global
// scope is accessible (via closure semantics) inside class methods that read global variables.
func TestASPScriptingDictionaryClassScope(t *testing.T) {
	source := `<%
Dim globalObj
Set globalObj = Server.CreateObject("Scripting.Dictionary")
globalObj.Add "Key", "Global Object Value"

Class TestClass
    Public Function ReadGlobal()
        If IsObject(globalObj) Then
            ReadGlobal = globalObj("Key")
        Else
            ReadGlobal = "not an object"
        End If
    End Function
End Class

Dim c
Set c = New TestClass
Response.Write c.ReadGlobal()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Global Object Value" {
		t.Fatalf("expected 'Global Object Value', got %q", output.String())
	}
}

// TestASPScriptingDictionaryForEach verifies For Each iteration over Scripting.Dictionary returns keys.
func TestASPScriptingDictionaryForEach(t *testing.T) {
	source := `<%
Dim d, k, keysStr
Set d = Server.CreateObject("Scripting.Dictionary")
d.Add "Key1", "Value1"
d.Add "Key2", "Value2"
d.Remove "Key1"
d.Add "Key3", "Value3"

keysStr = ""
For Each k In d
    keysStr = keysStr & k & ";"
Next

Response.Write keysStr
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	rendered := output.String()
	if rendered != "Key2;Key3;" {
		t.Fatalf("expected dictionary insertion-order output, got %q", rendered)
	}
}

// TestASPScriptingDictionaryForEachObjectMemberByNumericKey verifies
// dictionary-key iteration with numeric keys still supports chained object
// member access via dict(key).Member, matching Classic ASP breadcrumb/menu patterns.
func TestASPScriptingDictionaryForEachObjectMemberByNumericKey(t *testing.T) {
	source := `<%
Class PageObj
	Public title
	Public Property Get getParentLink
		getParentLink = title
	End Property
End Class

Dim d, k, p, out
Set d = Server.CreateObject("Scripting.Dictionary")

Set p = New PageObj
p.title = "A"
d.Add 10, p

Set p = New PageObj
p.title = "B"
d.Add 20, p

out = ""
For Each k In d
	out = out & d(k).getParentLink & " / "
Next

Response.Write out
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "A / B / " {
		t.Fatalf("unexpected dictionary object-member output: %q", output.String())
	}
}

// TestASPClassDefaultFunctionCall verifies object(...) dispatches to Public Default Function.
func TestASPClassDefaultFunctionCall(t *testing.T) {
	source := `<%
Class TestClass
    Public Default Function MyDefault(x)
        MyDefault = "Hello " & x
    End Function
End Class

Dim t
Set t = New TestClass
Response.Write t("World")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Hello World" {
		t.Fatalf("expected 'Hello World', got %q", output.String())
	}
}

// TestASPClassDefaultFunctionObjectResultChain verifies default function calls
// can return objects that are immediately consumed by IsObject and chained
// member calls without triggering default-property assignment errors.
func TestASPClassDefaultFunctionObjectResultChain(t *testing.T) {
	source := `<%
Class Child
	Public Function Serialize()
		Serialize = "ok"
	End Function
End Class

Class Pair
	Private i_value

	Public Property Get Value
		If IsObject(i_value) Then
			Set Value = i_value
		Else
			Value = i_value
		End If
	End Property

	Public Property Let Value(v)
		i_value = v
	End Property

	Public Property Set Value(v)
		Set i_value = v
	End Property
End Class

Class Container
	Private i_defaultPropertyName
	Private i_pair

	Private Sub Class_Initialize()
		i_defaultPropertyName = "data"
		Set i_pair = New Pair
		Set i_pair.Value = New Child
	End Sub

	Public Property Get defaultPropertyName
		defaultPropertyName = i_defaultPropertyName
	End Property

	Public Default Function Value(prop)
		If IsObject(i_pair.Value) Then
			Set Value = i_pair.Value
		Else
			Value = i_pair.Value
		End If
	End Function
End Class

Dim c
Set c = New Container
If IsObject(c(c.defaultPropertyName)) Then
	Response.Write c(c.defaultPropertyName).Serialize()
Else
	Response.Write "not-object"
End If
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("unexpected chained default function output: got %q want %q", output.String(), "ok")
	}
}

// TestASPErrObjectPropertyRoundTrip verifies intrinsic Err property read/write and Clear behavior.
func TestASPErrObjectPropertyRoundTrip(t *testing.T) {
	source := `<%
On Error Resume Next
Err.Clear
Err.ASPCode = "100"
Err.Category = "Syntax Error"
Err.File = "test_issues.asp"
Err.Line = 42
Err.Column = 15
Err.ASPDescription = "Extended ASP error description"
Err.Number = 11
Err.Description = "Division by zero"
Err.Source = "Test Source"

Response.Write Err.ASPCode & "|" & Err.Category & "|" & Err.File & "|" & Err.Line & "|" & Err.Column & "|" & Err.ASPDescription & "|" & Err.Number & "|" & Err.Description & "|" & Err.Source
Err.Clear
Response.Write "|" & Err.Number & "|" & Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "100|Syntax Error|test_issues.asp|42|15|Extended ASP error description|11|Division by zero|Test Source|0|"
	if output.String() != expected {
		t.Fatalf("unexpected Err output: got %q want %q", output.String(), expected)
	}
}

// TestASPErrRaiseResumeNextPopulatesErr verifies Err.Raise populates Err fields
// and execution continues under On Error Resume Next.
func TestASPErrRaiseResumeNextPopulatesErr(t *testing.T) {
	source := `<%
On Error Resume Next
Err.Clear
Err.Raise 100, "TestSource", "This is a test error"
Response.Write Err.Number & "|" & Err.Source & "|" & Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "100|TestSource|This is a test error"
	if output.String() != expected {
		t.Fatalf("unexpected Err.Raise output: got %q want %q", output.String(), expected)
	}
}

// TestASPErrRaiseResumeNextPopulatesHelpFields verifies Err.Raise optional
// helpfile and helpcontext arguments populate intrinsic Err properties.
func TestASPErrRaiseResumeNextPopulatesHelpFields(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.onResumeNext = true

	_ = vm.errRaise([]Value{
		NewInteger(101),
		NewString("HelpSource"),
		NewString("Help error"),
		NewString("help.chm"),
		NewInteger(42),
	})
	if vm.errObject == nil {
		t.Fatal("expected intrinsic Err object to be initialized")
	}
	if vm.errObject.HelpFile != "help.chm" || vm.errObject.HelpContext != 42 {
		t.Fatalf("unexpected internal Err object help fields: helpFile=%q helpContext=%d", vm.errObject.HelpFile, vm.errObject.HelpContext)
	}

	if got := vm.errPropertyValue("Number").Num; got != 101 {
		t.Fatalf("unexpected Err.Number: got %d want %d", got, 101)
	}
	if got := vm.errPropertyValue("HelpFile").String(); got != "help.chm" {
		t.Fatalf("unexpected Err.HelpFile: got %q want %q", got, "help.chm")
	}
	if got := vm.errPropertyValue("HelpContext").Num; got != 42 {
		t.Fatalf("unexpected Err.HelpContext: got %d want %d", got, 42)
	}
}

// TestASPErlReturnsLastErrorLine verifies Erl mirrors Err.Line for the most recent runtime error.
func TestASPErlReturnsLastErrorLine(t *testing.T) {
	source := `<%
On Error Resume Next
Dim x
x = 1 / 0
Response.Write Erl & "|" & Err.Line
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	parts := strings.Split(output.String(), "|")
	if len(parts) != 2 {
		t.Fatalf("expected Erl|Err.Line output, got %q", output.String())
	}
	if parts[0] == "0" || parts[0] == "" {
		t.Fatalf("expected non-zero Erl output, got %q", output.String())
	}
	if parts[0] != parts[1] {
		t.Fatalf("expected Erl to match Err.Line, got %q", output.String())
	}
}

// TestASPErrRaiseWithoutResumeNextHalts verifies Err.Raise propagates as runtime
// error when On Error Resume Next is not active.
func TestASPErrRaiseWithoutResumeNextHalts(t *testing.T) {
	source := `<%
On Error GoTo 0
Err.Clear
Err.Raise 200, "TestSource2", "Hard raise"
Response.Write "after"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	err := vm.Run()
	if err == nil {
		t.Fatal("expected Err.Raise to halt execution, got nil error")
	}

	vbErr, ok := err.(*VMError)
	if !ok {
		t.Fatalf("expected *VMError, got %T (%v)", err, err)
	}
	if vbErr.Number != 200 {
		t.Fatalf("unexpected raised number: got %d want %d", vbErr.Number, 200)
	}
	if vbErr.Source != "TestSource2" {
		t.Fatalf("unexpected raised source: got %q want %q", vbErr.Source, "TestSource2")
	}
	if vbErr.Description != "Hard raise" {
		t.Fatalf("unexpected raised description: got %q want %q", vbErr.Description, "Hard raise")
	}
	if output.String() != "" {
		t.Fatalf("expected no output after hard raise, got %q", output.String())
	}
}

func TestASPOnErrorGoto0ClearsErrState(t *testing.T) {
	source := `<%
On Error Resume Next
Dim x
x = 1 / 0
Response.Write Err.Number
Response.Write "|"
On Error GoTo 0
Response.Write Err.Number
Response.Write "|"
Response.Write Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	parts := strings.Split(output.String(), "|")
	if len(parts) != 3 {
		t.Fatalf("expected Err state triplet, got %q", output.String())
	}
	if parts[0] == "0" || parts[0] == "" {
		t.Fatalf("expected first Err.Number to be non-zero, got %q", parts[0])
	}
	if parts[1] != "0" {
		t.Fatalf("expected Err.Number reset to 0 after On Error GoTo 0, got %q", parts[1])
	}
	if parts[2] != "" {
		t.Fatalf("expected Err.Description reset after On Error GoTo 0, got %q", parts[2])
	}
}

// TestASPJSONobjectDuplicateDataRaiseDoesNotAbort verifies compatibility for
// legacy JSONobject scripts that may raise duplicate "data" key conflicts.
func TestASPJSONobjectDuplicateDataRaiseDoesNotAbort(t *testing.T) {
	source := `<%
Response.Write "A|"
Err.Raise 2, "JSONobject", "A property already exists with the name: data."
Response.Write Err.Number & "|" & Err.Description & "|B"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "A|0||B" {
		t.Fatalf("unexpected JSONobject compatibility output: got %q want %q", output.String(), "A|0||B")
	}
}

// TestASPClassFieldLetOverwritesObjectValue verifies class field Let assignment
// overwrites previous object values instead of requiring default property dispatch.
func TestASPClassFieldLetOverwritesObjectValue(t *testing.T) {
	source := `<%
Class Holder
	Private m_val
	Public Property Let V(x)
		m_val = x
	End Property
	Public Property Set V(x)
		Set m_val = x
	End Property
	Public Property Get Kind()
		Kind = TypeName(m_val)
	End Property
End Class

On Error Resume Next
Dim h, d
Set h = New Holder
Set d = CreateObject("Scripting.Dictionary")
Set h.V = d
h.V = 123
Response.Write h.Kind & "|" & Err.Number & "|" & Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "Integer|0|"
	if output.String() != expected {
		t.Fatalf("unexpected class field let overwrite output: got %q want %q", output.String(), expected)
	}
}

// TestASPPropertyGetReturnSlotOverwritesFromObjectToScalar verifies that
// assigning to a Property Get return slot overwrites directly instead of
// attempting default-property Let dispatch when the slot previously held object.
func TestASPPropertyGetReturnSlotOverwritesFromObjectToScalar(t *testing.T) {
	source := `<%
Class Holder
	Private m_val
	Public Property Let V(x)
		m_val = x
	End Property
	Public Property Set V(x)
		Set m_val = x
	End Property
	Public Property Get V
		If IsObject(m_val) Then
			Set V = m_val
		Else
			V = m_val
		End If
	End Property
End Class

On Error Resume Next
Dim h, d, tmp
Set h = New Holder
Set d = CreateObject("Scripting.Dictionary")
Set h.V = d
Set tmp = h.V
h.V = 123
Response.Write h.V & "|" & Err.Number & "|" & Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "123|0|"
	if output.String() != expected {
		t.Fatalf("unexpected property return overwrite output: got %q want %q", output.String(), expected)
	}
}

// TestASPLocalVariableOverwritesObjectWithScalar verifies plain local-variable
// assignment replaces an object reference instead of dispatching a default
// property write on the previous object value.
func TestASPLocalVariableOverwritesObjectWithScalar(t *testing.T) {
	source := `<%
Class Child
End Class

Function Render()
	Dim value
	Set value = New Child
	value = "ok"
	Render = value
End Function

On Error Resume Next
Response.Write Render() & "|" & Err.Number & "|" & Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "ok|0|"
	if output.String() != expected {
		t.Fatalf("unexpected local overwrite output: got %q want %q", output.String(), expected)
	}
}

// TestASPGlobalVariableOverwritesObjectWithScalar verifies plain global-variable
// assignment replaces an object reference instead of dispatching a default
// property write on the previous object value.
func TestASPGlobalVariableOverwritesObjectWithScalar(t *testing.T) {
	source := `<%
Class Child
End Class

On Error Resume Next
Dim value
Set value = New Child
value = "ok"
Response.Write value & "|" & Err.Number & "|" & Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "ok|0|"
	if output.String() != expected {
		t.Fatalf("unexpected global overwrite output: got %q want %q", output.String(), expected)
	}
}

// TestASPClassImplicitSetPropertyAssignment verifies Set assignments to class
// properties without explicit Me. are resolved as member property-set calls.
func TestASPClassImplicitSetPropertyAssignment(t *testing.T) {
	source := `<%
Class Box
	Private m
	Public Property Get Value
		If IsObject(m) Then
			Set Value = m
		Else
			Value = m
		End If
	End Property
	Public Property Set Value(v)
		Set m = v
	End Property
	Private Sub Class_Terminate
		If IsObject(Value) Then Set Value = Nothing
	End Sub
End Class

On Error Resume Next
Dim b, d
Set b = New Box
Set d = CreateObject("Scripting.Dictionary")
Set b.Value = d
Set b = Nothing
Response.Write "OK|" & Err.Number & "|" & Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "OK|0|"
	if output.String() != expected {
		t.Fatalf("unexpected implicit Set property output: got %q want %q", output.String(), expected)
	}
}

// TestASPExecuteGlobalDefinesFunctions verifies ExecuteGlobal can inject global
// function definitions that are callable later in the same ASP script.
func TestASPExecuteGlobalDefinesFunctions(t *testing.T) {
	source := `<%
Dim code1, code2, code3, q

code1 = "Function Test1()" & vbCrLf & _
        "Test1 = ""result1""" & vbCrLf & _
        "End Function"
ExecuteGlobal code1

code2 = "Function Test2(x)" & vbCrLf & _
        "Test2 = x" & vbCrLf & _
        "End Function"
ExecuteGlobal code2

q = Chr(34)
code3 = "Function Test3(x)" & vbCrLf & _
        "Test3 = x & " & q & " processed" & q & vbCrLf & _
        "End Function"
ExecuteGlobal code3

	Response.Write Test1() & "|" & Test2("hello") & "|" & Test3("hello")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "result1|hello|hello processed"
	if output.String() != expected {
		t.Fatalf("unexpected ExecuteGlobal output: got %q want %q", output.String(), expected)
	}
}

// TestASPExecuteGlobalResumeNextCompileError verifies ExecuteGlobal syntax errors
// populate Err instead of terminating when On Error Resume Next is active.
func TestASPExecuteGlobalResumeNextCompileError(t *testing.T) {
	source := `<%
On Error Resume Next
Err.Clear
ExecuteGlobal "Function Broken(" 
Response.Write Err.Number & "|" & Err.Source & "|" & Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expectedNumber := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.ExpectedEnd))
	expected := expectedNumber + "|VBScript compilation error|Missing 'End'"
	if output.String() != expected {
		t.Fatalf("unexpected ExecuteGlobal compile error output: got %q want %q", output.String(), expected)
	}
}

// TestASPExecuteRunsDynamicCode verifies Execute evaluates dynamic code in the active script scope.
func TestASPExecuteRunsDynamicCode(t *testing.T) {
	source := `<%
Dim x
x = 10
Execute "x = x + 5"
Response.Write x
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "15" {
		t.Fatalf("unexpected Execute output: got %q want %q", output.String(), "15")
	}
}

// TestASPEvalResolvesImplicitClassMemberInMethod verifies Eval can resolve
// implicit same-class members when called inside one class method.
func TestASPEvalResolvesImplicitClassMemberInMethod(t *testing.T) {
	source := `<%
Class Accumulator
	Private value

	Public Sub Class_Initialize()
		value = 41
	End Sub

	Public Function ReadWithEval()
		ReadWithEval = Eval("value")
	End Function
End Class

Dim o
Set o = New Accumulator
Response.Write o.ReadWithEval()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "41" {
		t.Fatalf("unexpected Eval implicit member output: got %q want %q", output.String(), "41")
	}
}

// TestASPExecuteResolvesImplicitClassMemberInMethod verifies Execute can read
// and write implicit same-class members inside one class method.
func TestASPExecuteResolvesImplicitClassMemberInMethod(t *testing.T) {
	source := `<%
Class Counter
	Private total

	Public Sub Class_Initialize()
		total = 5
	End Sub

	Public Function Step()
		Execute "total = total + 2"
		Step = total
	End Function
End Class

Dim c
Set c = New Counter
Response.Write c.Step()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "7" {
		t.Fatalf("unexpected Execute implicit member output: got %q want %q", output.String(), "7")
	}
}

// TestASPExecuteGlobalResolvesImplicitClassMemberInMethod verifies ExecuteGlobal
// can resolve implicit same-class members when called from a class method.
func TestASPExecuteGlobalResolvesImplicitClassMemberInMethod(t *testing.T) {
	source := `<%
Class Counter
	Private total

	Public Sub Class_Initialize()
		total = 9
	End Sub

	Public Function Step()
		ExecuteGlobal "total = total + 3"
		Step = total
	End Function
End Class

Dim c
Set c = New Counter
Response.Write c.Step()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	// ExecuteGlobal runs in global scope and cannot access the class-private 'total' member.
	// The assignment 'total = total + 3' creates a new global 'total' = 3, leaving the
	// class field unchanged at 9. Classic ASP/VBScript behavior: Step() returns 9.
	if output.String() != "9" {
		t.Fatalf("unexpected ExecuteGlobal implicit member output: got %q want %q", output.String(), "9")
	}
}

// TestASPExecuteGlobalSyncsDynamicNativeObjectMaps verifies native object maps
// created during ExecuteGlobal are available in the parent VM after sync.
func TestASPExecuteGlobalSyncsDynamicNativeObjectMaps(t *testing.T) {
	source := `<%
ExecuteGlobal "Set d = Server.CreateObject(""Scripting.Dictionary"") : d.Add ""k"", ""v"""
Response.Write d.Item("k")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "v" {
		t.Fatalf("unexpected ExecuteGlobal native map sync output: got %q want %q", output.String(), "v")
	}
}

// TestASPExecuteGlobalCanInstantiateParentClass verifies dynamic code executed via
// ExecuteGlobal can instantiate one class declared in the parent ASP script.
func TestASPExecuteGlobalCanInstantiateParentClass(t *testing.T) {
	source := `<%
Class MainClass
	Public Function Name()
		Name = "main"
	End Function
End Class

Dim dynamicCode
dynamicCode = "Dim o" & vbCrLf & _
	          "Set o = New MainClass" & vbCrLf & _
	          "Response.Write o.Name()"
ExecuteGlobal dynamicCode
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "main" {
		t.Fatalf("unexpected ExecuteGlobal parent-class output: got %q want %q", output.String(), "main")
	}
}

// TestASPExecuteGlobalCanUseClassDeclaredAfterCall verifies global class declarations
// remain available to ExecuteGlobal even when the call appears before Class...End Class.
func TestASPExecuteGlobalCanUseClassDeclaredAfterCall(t *testing.T) {
	source := `<%
Dim dynamicCode
dynamicCode = "Dim o" & vbCrLf & _
	          "Set o = New DelayedClass" & vbCrLf & _
	          "Response.Write o.Name()"
ExecuteGlobal dynamicCode

Class DelayedClass
	Public Function Name()
		Name = "delayed"
	End Function
End Class
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "delayed" {
		t.Fatalf("unexpected ExecuteGlobal delayed-class output: got %q want %q", output.String(), "delayed")
	}
}

// TestASPExecuteGlobalSeesClassDeclaredByPreviousExecuteGlobal verifies one
// later ExecuteGlobal block can instantiate a class declared in a prior dynamic block.
func TestASPExecuteGlobalSeesClassDeclaredByPreviousExecuteGlobal(t *testing.T) {
	source := `<%
Dim defCode, runCode
defCode = "Class DynClass" & vbCrLf & _
	         "  Public Function Name()" & vbCrLf & _
	         "    Name = ""dyn""" & vbCrLf & _
	         "  End Function" & vbCrLf & _
	         "End Class"
ExecuteGlobal defCode

runCode = "Dim o" & vbCrLf & _
	         "Set o = New DynClass" & vbCrLf & _
	         "Response.Write o.Name()"
ExecuteGlobal runCode
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "dyn" {
		t.Fatalf("unexpected ExecuteGlobal dynamic-class output: got %q want %q", output.String(), "dyn")
	}
}

// TestASPEvalNewSeesParentClassAfterExecuteGlobal verifies Eval("New ...") can
// instantiate one parent class after ExecuteGlobal has run in the same request.
func TestASPEvalNewSeesParentClassAfterExecuteGlobal(t *testing.T) {
	source := `<%
Class MainClass
	Public Function Name()
		Name = "main"
	End Function
End Class

ExecuteGlobal "Dim marker : marker = 1"
Dim o
Set o = Eval("New MainClass")
Response.Write o.Name()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "main" {
		t.Fatalf("unexpected Eval New parent-class output: got %q want %q", output.String(), "main")
	}
}

// TestASPExecuteGlobalInsideClassMethodRegistersClassGlobally verifies
// ExecuteGlobal called inside a class method still defines classes globally,
// allowing a later Eval("New ...") call to instantiate that class.
func TestASPExecuteGlobalInsideClassMethodRegistersClassGlobally(t *testing.T) {
	source := `<%
Class Host
	Public Sub LoadType()
		ExecuteGlobal "Class DynType" & vbCrLf & _
		              "  Public Function Name()" & vbCrLf & _
		              "    Name = ""dyn""" & vbCrLf & _
		              "  End Function" & vbCrLf & _
		              "End Class"
	End Sub

	Public Function Build()
		Set Build = Eval("new DynType")
	End Function
End Class

Dim h, o
Set h = New Host
h.LoadType
Set o = h.Build
Response.Write o.Name()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "dyn" {
		t.Fatalf("unexpected class-method ExecuteGlobal output: got %q want %q", output.String(), "dyn")
	}
}

// TestASPEvalNewWithConcatenatedClassName verifies Eval can create objects when
// the class name is built dynamically (e.g. "new cls_" & value).
func TestASPEvalNewWithConcatenatedClassName(t *testing.T) {
	source := `<%
Class cls_asplite_json
	Public Function Name()
		Name = "json"
	End Function
End Class

Class Loader
	Public Function plugin(value)
		Set plugin = Eval("new cls_asplite_" & value)
	End Function
End Class

Dim l, o
Set l = New Loader
Set o = l.plugin("json")
Response.Write o.Name()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "json" {
		t.Fatalf("unexpected Eval concatenated New output: got %q want %q", output.String(), "json")
	}
}

// TestASPGetRefCallsUserFunction verifies GetRef can return and invoke a global function reference.
func TestASPGetRefCallsUserFunction(t *testing.T) {
	source := `<%
Function AddOne(x)
	AddOne = x + 1
End Function

Dim refFn
Set refFn = GetRef("AddOne")
Response.Write refFn(9)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "10" {
		t.Fatalf("unexpected GetRef output: got %q want %q", output.String(), "10")
	}
}

// TestASPCIntStringConversionAndResumeNext verifies valid string conversion and invalid-string Err handling.
func TestASPCIntStringConversionAndResumeNext(t *testing.T) {
	source := `<%
On Error Resume Next
Response.Write CInt("30")
Response.Write "|"
Response.Write Err.Number
Dim badValue
badValue = CInt("not-a-number")
Response.Write "|"
Response.Write Err.Number
Response.Write "|"
Response.Write Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "30|0|" + strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.TypeMismatch)) + "|Type mismatch"
	if output.String() != expected {
		t.Fatalf("unexpected CInt output: got %q want %q", output.String(), expected)
	}
}

// TestASPNumericConversionsRaiseVBScriptErrors verifies string conversion failures
// use VBScript Err semantics and that CInt accepts hex string values.
func TestASPNumericConversionsRaiseVBScriptErrors(t *testing.T) {
	source := `<%
On Error Resume Next
Dim value

value = CLng("not-a-number")
Response.Write Err.Number & "|" & Err.Description
Err.Clear
Response.Write "|"

value = CDbl("xyz123")
Response.Write Err.Number & "|" & Err.Description
Err.Clear
Response.Write "|"

value = CSng("abc")
Response.Write Err.Number & "|" & Err.Description
Err.Clear
Response.Write "|"

value = CByte("invalid")
Response.Write Err.Number & "|" & Err.Description
Err.Clear
Response.Write "|"

Response.Write CInt("&HFF") & "|" & Err.Number & "|" & Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	typeMismatch := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.TypeMismatch))
	expected := typeMismatch + "|Type mismatch|" +
		typeMismatch + "|Type mismatch|" +
		typeMismatch + "|Type mismatch|" +
		typeMismatch + "|Type mismatch|" +
		"255|0|"
	if output.String() != expected {
		t.Fatalf("unexpected numeric conversion output: got %q want %q", output.String(), expected)
	}
}

// TestASPCByteOverflowRaisesError verifies CByte out-of-range values raise VBScript Overflow.
func TestASPCByteOverflowRaisesError(t *testing.T) {
	source := `<%
On Error Resume Next
Dim value
value = CByte(256)
Response.Write Err.Number & "|" & Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.Overflow)) + "|Overflow"
	if output.String() != expected {
		t.Fatalf("unexpected CByte overflow output: got %q want %q", output.String(), expected)
	}
}

// TestASPOptionCompareTextAndEval verifies Option Compare Text and Eval arithmetic semantics.
func TestASPOptionCompareTextAndEval(t *testing.T) {
	source := `<%
Option Compare Text
If "ABC" = "abc" Then
    Response.Write "eq"
Else
    Response.Write "neq"
End If
Response.Write "|"
Response.Write Eval("10 + 20")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "eq|30" {
		t.Fatalf("unexpected Option Compare/Eval output: got %q want %q", output.String(), "eq|30")
	}
}

// TestASPOptionCompareBinaryInAspBlock verifies Option Compare Binary compiles from ASP blocks without assignment parser fallback.
func TestASPOptionCompareBinaryInAspBlock(t *testing.T) {
	source := `<%
Option Compare Binary
If ("ABC" = "abc") = False Then
    Response.Write "ok"
Else
    Response.Write "bad"
End If
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Fatalf("unexpected Option Compare Binary output: %q", output.String())
	}
}

// TestASPEvalStringLiteralUsesCallerScope verifies Eval with a string literal
// resolves variables, arrays, and built-in function calls in the current scope.
func TestASPEvalStringLiteralUsesCallerScope(t *testing.T) {
	source := `<%
Dim x, y, z
Dim arr(1)

x = 100
y = 200
z = 300
arr(0) = "apple"
arr(1) = "banana"

Response.Write Eval("x") & "|"
Response.Write Eval("x + 5") & "|"
Response.Write Eval("(x + y) * z") & "|"
Response.Write Eval("Len(""hello"")") & "|"
Response.Write Eval("arr(0)") & "|"
Response.Write Eval("arr(1)")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "100|105|90000|5|apple|banana"
	if output.String() != expected {
		t.Fatalf("unexpected Eval caller-scope output: got %q want %q", output.String(), expected)
	}
}

// TestASPEvalVariableExpressionRuntimePath verifies Eval with a non-literal
// argument compiles and executes through the runtime Eval path.
func TestASPEvalVariableExpressionRuntimePath(t *testing.T) {
	source := `<%
Dim expr
expr = "true"
Response.Write TypeName(Eval(expr))
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Boolean" {
		t.Fatalf("unexpected Eval runtime-path output: got %q want %q", output.String(), "Boolean")
	}
}

// TestASPCurrencyDecimalConversionSemantics verifies CCur/CDec share numeric coercion and error propagation.
func TestASPCurrencyDecimalConversionSemantics(t *testing.T) {
	source := `<%
On Error Resume Next
Response.Write CCur("&H10") & "|" & CDec("&O10") & "|"
Dim value
value = CCur("invalid")
Response.Write Err.Number & "|" & Err.Description & "|"
Err.Clear
value = CDec("invalid")
Response.Write Err.Number & "|" & Err.Description
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	typeMismatch := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.TypeMismatch))
	expected := "16|8|" + typeMismatch + "|Type mismatch|" + typeMismatch + "|Type mismatch"
	if output.String() != expected {
		t.Fatalf("unexpected CCur/CDec output: got %q want %q", output.String(), expected)
	}
}

// TestASPIsOperatorObjectReferenceSemantics verifies Is/Is Not for object references and null values.
func TestASPIsOperatorObjectReferenceSemantics(t *testing.T) {
	source := `<%
On Error Resume Next
Dim objShell
Set objShell = CreateObject("WScript.Shell")
Response.Write (objShell Is Nothing)
Response.Write "|"
Response.Write (objShell Is Not Nothing)
Response.Write "|"

Dim emptyVal
Response.Write (emptyVal Is Nothing)
Response.Write "|" & Err.Number & "|" & Err.Description
Err.Clear
Response.Write "|"

Dim nullVal
nullVal = Null
Response.Write (nullVal Is Nothing)
Response.Write "|" & Err.Number & "|" & Err.Description
Err.Clear
Response.Write "|"

Dim json
Dim jsonObj
Set json = CreateObject("G3JSON")
Set jsonObj = json.NewObject()
Response.Write (jsonObj Is Nothing)
Response.Write "|"
Response.Write (jsonObj Is Not Nothing)
Response.Write "|" & Err.Number & "|" & Err.Description
Err.Clear
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	typeMismatch := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.TypeMismatch))
	expected := "False|True||" + typeMismatch + "|Type mismatch||" + typeMismatch + "|Type mismatch|False|True|0|"
	if output.String() != expected {
		t.Fatalf("unexpected Is/Is Not output: got %q want %q", output.String(), expected)
	}
}

// TestASPWScriptShellExpandEnvironmentStrings verifies ASP runtime compatibility
// for CreateObject("WScript.Shell").ExpandEnvironmentStrings("%NAME%").
func TestASPWScriptShellExpandEnvironmentStrings(t *testing.T) {
	const envName = "AXONASP_WSCRIPT_ASP_EXPAND_TEST"
	const envValue = "asp-expand-ok"
	if err := os.Setenv(envName, envValue); err != nil {
		t.Fatalf("setenv failed: %v", err)
	}
	t.Cleanup(func() {
		_ = os.Unsetenv(envName)
	})

	source := `<%
Dim wshell
Dim dbqString
Set wshell = CreateObject("WScript.Shell")
dbqString = wshell.ExpandEnvironmentStrings("%AXONASP_WSCRIPT_ASP_EXPAND_TEST%")
Response.Write dbqString
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != envValue {
		t.Fatalf("unexpected expand output: got %q want %q", output.String(), envValue)
	}
}

// TestASPNotIsNothingNativeObjectSemantics verifies VBScript idiom:
// "If Not obj Is Nothing Then" must compile as Not (obj Is Nothing).
func TestASPNotIsNothingNativeObjectSemantics(t *testing.T) {
	source := `<%
On Error Resume Next
Dim zip
Set zip = CreateObject("G3ZIP")
If Not zip Is Nothing Then
	Response.Write "OK"
Else
	Response.Write "NO"
End If
Response.Write "|" & Err.Number & "|" & Err.Description
Err.Clear
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "OK|0|"
	if output.String() != expected {
		t.Fatalf("unexpected Not Is Nothing output: got %q want %q", output.String(), expected)
	}
}

// TestASPNotTypeMismatchNonNumericString verifies that applying Not to a non-numeric
// string raises Type mismatch (error 13), consistent with original VBScript behaviour
// where Not requires a CLng-compatible operand.
func TestASPNotTypeMismatchNonNumericString(t *testing.T) {
	source := `<%
On Error Resume Next
Dim result
result = Not "hello"
Response.Write result & "|" & Err.Number
Err.Clear
result = Not "True"
Response.Write "|" & result & "|" & Err.Number
Err.Clear
result = Not "123"
Response.Write "|" & result
Err.Clear
result = Not "0"
Response.Write "|" & result
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	typeMismatch := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.TypeMismatch))
	// Not "hello" → Type mismatch, result stays Empty ("").
	// Not "True" → Type mismatch, result stays Empty ("").
	// Not "123" → Not 123 = -124.
	// Not "0" → Not 0 = -1.
	expected := "|" + typeMismatch + "||" + typeMismatch + "|-124|-1"
	if output.String() != expected {
		t.Fatalf("unexpected Not string output: got %q want %q", output.String(), expected)
	}
}

// TestASPRegExpSubMatchesValueProperty verifies SubMatches.Item(i).Value returns each capture text.
func TestASPRegExpSubMatchesValueProperty(t *testing.T) {
	source := `<%
Dim regex, matches, firstMatch, subMatches
Set regex = New RegExp
regex.Pattern = "(\d{3})-(\w+)-([a-z]+)"
regex.Global = False
regex.IgnoreCase = True

Set matches = regex.Execute("123-TEST-abc")
Set firstMatch = matches.Item(0)
Set subMatches = firstMatch.SubMatches

Response.Write subMatches.Item(0).Value & "|" & subMatches.Item(1).Value & "|" & subMatches.Item(2).Value
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "123|TEST|abc" {
		t.Fatalf("unexpected SubMatches output: %q", output.String())
	}
}

// TestASPErrObjectASPCodeStringRoundTrip verifies Err.ASPCode preserves alphanumeric assignments.
func TestASPErrObjectASPCodeStringRoundTrip(t *testing.T) {
	source := `<%
On Error Resume Next
Err.Clear
Err.ASPCode = "ERR_TEST_001"
Response.Write Err.ASPCode
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ERR_TEST_001" {
		t.Fatalf("unexpected Err.ASPCode output: %q", output.String())
	}
}

// TestASPResponseWriteNowWithoutParens verifies zero-arg builtin call semantics in expression position.
func TestASPResponseWriteNowWithoutParens(t *testing.T) {
	source := `<%
Response.Write Now
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	rendered := output.String()
	if rendered == "" || strings.Contains(rendered, "[Builtin:") {
		t.Fatalf("unexpected Response.Write Now output: %q", rendered)
	}
}

// TestASPNestedBuiltinNowWithoutParens verifies implicit call also works when used as another builtin argument.
func TestASPNestedBuiltinNowWithoutParens(t *testing.T) {
	source := `<%
Response.Write Left(Now, 1)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	rendered := output.String()
	if rendered == "[" || strings.Contains(rendered, "[Builtin:") {
		t.Fatalf("unexpected nested builtin output: %q", rendered)
	}
}

// TestASPClassSetAssignmentFunctionWithoutParens verifies Set x = functionName resolves implicit self call.
func TestASPClassSetAssignmentFunctionWithoutParens(t *testing.T) {
	source := `<%
Class DictProvider
    Public Function GetDict()
        Set GetDict = Server.CreateObject("Scripting.Dictionary")
        GetDict.Add "key1", "val1"
    End Function

    Public Sub Run()
        Dim d
        Set d = GetDict
        Response.Write d.Count
    End Sub
End Class

Dim dp
Set dp = New DictProvider
dp.Run
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "1" {
		t.Fatalf("unexpected implicit Set-call output: %q", output.String())
	}
}

// TestASPClassFieldDictionaryIndexedAccess verifies a Dictionary stored in a class
// field preserves default-member calls when accessed with arguments inside class code.
func TestASPClassFieldDictionaryIndexedAccess(t *testing.T) {
	source := `<%
Class FieldHolder
	Private allFields

	Private Sub Class_Initialize()
		Set allFields = CreateObject("Scripting.Dictionary")
		Dim item
		Set item = CreateObject("Scripting.Dictionary")
		item.Add "type", "hidden"
		allFields.Add 0, item
	End Sub

	Public Sub Run()
		Response.Write allFields(0)("type")
	End Sub
End Class

Dim holder
Set holder = New FieldHolder
holder.Run
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "hidden" {
		t.Fatalf("unexpected class-field Dictionary indexed access output: %q", output.String())
	}
}

// TestASPClassFieldDictionaryArrayObjectAssignment verifies Set arr(k) = dict(k)
// works when the Dictionary lives in a class field and stores object values.
func TestASPClassFieldDictionaryArrayObjectAssignment(t *testing.T) {
	source := `<%
Class FB
	Private allFields, counter

	Private Sub Class_Initialize()
		Set allFields = CreateObject("Scripting.Dictionary")
		counter = 0
		Dim h
		Set h = field("hidden")
		h.Add "name", "sys"
	End Sub

	Public Function field(value)
		Set field = CreateObject("Scripting.Dictionary")
		field.Add "type", value
		allFields.Add counter, field
		counter = counter + 1
	End Function

	Public Sub build()
		Dim arr : arr = Array()
		ReDim arr(counter-1)
		Dim k
		For Each k In allFields
			Set arr(k) = allFields(k)
		Next
		Response.Write TypeName(arr(0)) & "|" & arr(0)("type")
	End Sub
End Class

Dim form
Set form = New FB
Dim f
Set f = form.field("text")
form.build
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "Dictionary|hidden" {
		t.Fatalf("unexpected class-field Dictionary array object assignment output: %q", output.String())
	}
}

// TestASPNestedDictionaryRequestFormAssignment verifies the aspLite-style
// pattern allFields(fieldKey)("value") = Request.Form(allFields(fieldKey)("name"))
// preserves all posted values for a repeated checkbox field.
func TestASPNestedDictionaryRequestFormAssignment(t *testing.T) {
	source := `<%
Class FB
	Private allFields, counter

	Private Sub Class_Initialize()
		Set allFields = CreateObject("Scripting.Dictionary")
		counter = 0
	End Sub

	Public Function field(value)
		Set field = CreateObject("Scripting.Dictionary")
		field.Add "type", value
		allFields.Add counter, field
		counter = counter + 1
	End Function

	Public Sub build()
		Dim fieldkey
		For Each fieldkey In allFields
			If allFields(fieldkey).Exists("name") Then
				If Not IsEmpty(Request.Form(allFields(fieldkey)("name"))) Then
					allFields(fieldkey)("value") = Request.Form(allFields(fieldkey)("name"))
				End If
			End If
		Next
		Response.Write allFields(0)("value")
	End Sub
End Class

Dim form, checkboxes
Set form = New FB
Set checkboxes = form.field("checkbox")
checkboxes.Add "name", "checkboxes"
form.build
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Request().Form.AddValues("checkboxes", []string{"1", "2"})
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "1, 2" {
		t.Fatalf("expected dictionary count output 1, got %q", output.String())
	}
}

// TestASPGlobalSetAssignmentZeroArgFunctionWithoutParens verifies Set x = funcName
// invokes a page-scope zero-argument Function and stores its returned object reference.
func TestASPGlobalSetAssignmentZeroArgFunctionWithoutParens(t *testing.T) {
	source := `<%
Function GetDict()
	Set GetDict = Server.CreateObject("Scripting.Dictionary")
	GetDict.Add "key1", "val1"
End Function

Dim d
Set d = GetDict
Response.Write d.Count
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "1" {
		t.Fatalf("expected dictionary count output 1, got %q", output.String())
	}
}

// TestASPNestedDictionaryRequestFormReadArgument verifies nested Dictionary reads
// are preserved when used as the argument expression for Request.Form(...).
func TestASPNestedDictionaryRequestFormReadArgument(t *testing.T) {
	source := `<%
Dim allFields, item
Set allFields = CreateObject("Scripting.Dictionary")
Set item = CreateObject("Scripting.Dictionary")
item.Add "name", "checkboxes"
allFields.Add 0, item

Response.Write allFields(0)("name") & "|"
Response.Write IsEmpty(Request.Form(allFields(0)("name"))) & "|"
Response.Write Request.Form(allFields(0)("name"))
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Request().Form.AddValues("checkboxes", []string{"1", "2"})
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "checkboxes|False|1, 2" {
		t.Fatalf("unexpected nested dictionary Request.Form argument output: %q", output.String())
	}
}

// TestASPNestedDictionaryWrite verifies default-member writes on a nested
// Scripting.Dictionary preserve the assigned scalar value.
func TestASPNestedDictionaryWrite(t *testing.T) {
	source := `<%
Dim allFields, item
Set allFields = CreateObject("Scripting.Dictionary")
Set item = CreateObject("Scripting.Dictionary")
allFields.Add 0, item

allFields(0)("value") = "1, 2"
Response.Write allFields(0)("value")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "1, 2" {
		t.Fatalf("unexpected nested dictionary write output: %q", output.String())
	}
}

// TestASPByRefRecursiveParenthesizedCall verifies ByRef propagation through recursive calls using parentheses.
func TestASPByRefRecursiveParenthesizedCall(t *testing.T) {
	source := `<%
Function testByRef(ByRef s, depth)
    If depth > 0 Then
        testByRef = testByRef(s, depth - 1)
    Else
        s = "MODIFIED_BY_INNER"
        testByRef = s
    End If
End Function

Dim myStr
myStr = "ORIGINAL"
Dim result
result = testByRef(myStr, 2)

Response.Write result & "|" & myStr
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "MODIFIED_BY_INNER|MODIFIED_BY_INNER" {
		t.Fatalf("unexpected ByRef recursive output: %q", output.String())
	}
}

func buildNestedWriteCallHighIndexSource(padCount int) string {
	var source strings.Builder
	source.WriteString("<%\n")
	for i := 0; i < padCount; i++ {
		source.WriteString("Dim pad")
		source.WriteString(strconv.Itoa(i))
		source.WriteString("\n")
	}
	source.WriteString(`Class FB
	Public Function field(value)
		Set field = CreateObject("Scripting.Dictionary")
		field.Add "type", value
	End Function
End Class

Dim form
Set form = New FB

Dim f
Set f = form.field("text")

Response.Write f("type")
%>`)
	return source.String()
}

// TestASPNoParenMemberCallNestedCallHighGlobalIndex verifies undoTrailingCoerce does not trim
// operand bytes when the call target global index ends with the OpCoerceToValue byte value.
func TestASPNoParenMemberCallNestedCallHighGlobalIndex(t *testing.T) {
	var compiler *Compiler
	padCount := -1
	for i := 0; i < 512; i++ {
		candidate := NewASPCompiler(buildNestedWriteCallHighIndexSource(i))
		if err := candidate.Compile(); err != nil {
			t.Fatalf("compile failed during padding search at %d: %v", i, err)
		}
		globalIdx, ok := candidate.Globals.Get("f")
		if !ok {
			t.Fatal("expected global slot for f")
		}
		if globalIdx&0xFF == int(OpCoerceToValue) {
			compiler = candidate
			padCount = i
			break
		}
	}
	if compiler == nil {
		t.Fatalf("could not align f global index low byte to %d", OpCoerceToValue)
	}
	if padCount < 0 {
		t.Fatal("missing matched padding count")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed with pad count %d: %v", padCount, err)
	}
	host.Response().Flush()

	if output.String() != "text" {
		t.Fatalf("unexpected nested write output: %q", output.String())
	}
}

// TestASPByRefRecursiveWithApplicationArray verifies recursive ByRef replacement when cache is initialized in Application.
func TestASPByRefRecursiveWithApplicationArray(t *testing.T) {
	source := `<%
Function outerFunc(ByRef s, fill)
	If InStr(1, s, "[PLACEHOLDER:") <> 0 Then
        If IsEmpty(Application("arr_cached")) Then
            Dim arr(1, 2)
            arr(0, 0) = 1 : arr(1, 0) = "CODE1"
            arr(0, 1) = 2 : arr(1, 1) = "CODE2"
            arr(0, 2) = 3 : arr(1, 2) = "MYCODE"
            Application("arr_cached") = arr
            outerFunc = outerFunc(s, fill)
        End If
    End If

    If Not IsEmpty(Application("arr_cached")) Then
        Dim i
        For i = LBound(Application("arr_cached"), 2) To UBound(Application("arr_cached"), 2)
            Dim code
            code = Application("arr_cached")(1, i)
            If InStr(1, LCase(s), "[placeholder:" & LCase(code) & "]") <> 0 Then
                s = Replace(s, "[PLACEHOLDER:" & code & "]", "REPLACED_" & code, 1, -1, 1)
            End If
        Next
    End If
    outerFunc = s
End Function

Application("arr_cached") = Empty
Dim testStr
testStr = "Hello [PLACEHOLDER:MYCODE] World"
Response.Write outerFunc(testStr, True) & "|" & testStr
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "Hello REPLACED_MYCODE World|Hello REPLACED_MYCODE World"
	if output.String() != expected {
		t.Fatalf("unexpected recursive Application ByRef output: %q", output.String())
	}
}

// TestASPExecuteGlobalZeroArgObjectMemberAutoCall verifies ExecuteGlobal keeps IIS-compatible
// autocall behavior for zero-arg global functions returning objects used before member access.
func TestASPExecuteGlobalZeroArgObjectMemberAutoCall(t *testing.T) {
	source := `<%
Function MakeObj()
	Dim d
	Set d = Server.CreateObject("Scripting.Dictionary")
	d.Add "id", "42"
	Set MakeObj = d
End Function

ExecuteGlobal "Function DynTest() : DynTest = MakeObj.Item(""id"") : End Function"
Response.Write DynTest()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "42" {
		t.Fatalf("unexpected ExecuteGlobal autocall output: %q", output.String())
	}
}

// TestASPByRefClassMemberArrayResize verifies ByRef propagation for class member arrays.
func TestASPByRefClassMemberArrayResize(t *testing.T) {
	source := `<%
Class Holder
	Private arr
	Private count

	Private Sub Class_Initialize
		ReDim arr(-1)
		count = 0
	End Sub

	Private Sub EnsureCapacity(ByRef targetArr, ByVal nextCount)
		If UBound(targetArr) < nextCount - 1 Then
			ReDim Preserve targetArr(nextCount - 1)
		End If
	End Sub

	Public Sub Add(ByVal v)
		EnsureCapacity arr, count + 1
		arr(count) = v
		count = count + 1
	End Sub

	Public Function JoinAll()
		Dim i, s
		s = ""
		For i = 0 To count - 1
			If i > 0 Then s = s & ","
			s = s & CStr(arr(i))
		Next
		JoinAll = s
	End Function
End Class

Dim h
Set h = New Holder
h.Add "a"
h.Add True
Response.Write h.JoinAll()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "a,True" {
		t.Fatalf("unexpected class member ByRef output: %q", output.String())
	}
}

// TestASPExitFor verifies Classic ASP Exit For behavior in nested For...Next compilation.
func TestASPExitFor(t *testing.T) {
	source := `<%
Dim i
For i = 1 To 5
    If i = 3 Then
        Exit For
    End If
    Response.Write i
Next
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "12" {
		t.Fatalf("unexpected Exit For output: %q", output.String())
	}
}

// TestASPExitDo verifies Classic ASP Exit Do behavior in Do...Loop compilation.
func TestASPExitDo(t *testing.T) {
	source := `<%
Dim i
i = 0
Do While i < 5
    i = i + 1
    If i = 3 Then
        Exit Do
    End If
    Response.Write i
Loop
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "12" {
		t.Fatalf("unexpected Exit Do output: %q", output.String())
	}
}

// TestASPIncludeInitializesNestedArrays verifies include expansion and nested array access (FontMap-style).
func TestASPIncludeInitializesNestedArrays(t *testing.T) {
	tmpDir := t.TempDir()
	mainPath := filepath.Join(tmpDir, "main.asp")
	includePath := filepath.Join(tmpDir, "captcha.inc")

	includeSource := `<%
Dim FontMap(1)
FontMap(0) = Array("13", "A", "B")
FontMap(1) = Array("14", "C", "D")
%>`
	if err := os.WriteFile(includePath, []byte(includeSource), 0o600); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	mainSource := `<!--#include file="captcha.inc"-->
<%
Response.Write UBound(FontMap(0)) & "|" & FontMap(0)(0) & "|" & FontMap(0)(1) & "|" & FontMap(1)(0)
%>`
	if err := os.WriteFile(mainPath, []byte(mainSource), 0o600); err != nil {
		t.Fatalf("write main failed: %v", err)
	}

	compiler := NewASPCompiler(mainSource)
	compiler.SetSourceName(mainPath)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	rendered := strings.TrimSpace(output.String())
	if rendered != "2|13|A|14" {
		t.Fatalf("unexpected include nested-array output: %q", output.String())
	}
}

// TestASPCreateObjectSupportsMSXML2XMLHTTP verifies CreateObject recognizes MSXML2.XMLHTTP ProgID.
func TestASPCreateObjectSupportsMSXML2XMLHTTP(t *testing.T) {
	source := `<%
Dim http
Set http = Server.CreateObject("MSXML2.XMLHTTP")
If IsObject(http) Then
    Response.Write "OK"
Else
    Response.Write "NO"
End If
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if strings.TrimSpace(output.String()) != "OK" {
		t.Fatalf("unexpected CreateObject output: %q", output.String())
	}
}

// TestASPConstDeclaration verifies Const declarations compile and evaluate correctly.
func TestASPConstDeclaration(t *testing.T) {
	source := `<%
Const A = 10, B = "X"
Response.Write A & "|" & B
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if strings.TrimSpace(output.String()) != "10|X" {
		t.Fatalf("unexpected Const output: %q", output.String())
	}
}

// TestASPIncludeFileWithRelativeSourceName verifies include file path resolution when sourceName is workspace-relative.
func TestASPIncludeFileWithRelativeSourceName(t *testing.T) {
	tmpDir := t.TempDir()
	mainPath := filepath.Join(tmpDir, "www", "tests", "main.asp")
	includePath := filepath.Join(tmpDir, "www", "tests", "inc_const.inc")

	if err := os.MkdirAll(filepath.Dir(mainPath), 0o700); err != nil {
		t.Fatalf("mkdir main failed: %v", err)
	}

	includeSource := `<%
Const Z = 42
%>`
	if err := os.WriteFile(includePath, []byte(includeSource), 0o600); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	mainSource := `<!--#include file="inc_const.inc"-->
<%
Response.Write Z
%>`
	if err := os.WriteFile(mainPath, []byte(mainSource), 0o600); err != nil {
		t.Fatalf("write main failed: %v", err)
	}

	oldWD, err := os.Getwd()
	if err != nil {
		t.Fatalf("getwd failed: %v", err)
	}
	defer func() {
		_ = os.Chdir(oldWD)
	}()
	if err := os.Chdir(tmpDir); err != nil {
		t.Fatalf("chdir failed: %v", err)
	}

	compiler := NewASPCompiler(mainSource)
	compiler.SetSourceName(filepath.ToSlash(filepath.Join("www", "tests", "main.asp")))
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if strings.TrimSpace(output.String()) != "42" {
		t.Fatalf("unexpected include output: %q", output.String())
	}
}

// TestASPIncludeCompileErrorReportsOriginalFileAndLine verifies compile errors inside
// included files are mapped back to the include file path and original line.
func TestASPIncludeCompileErrorReportsOriginalFileAndLine(t *testing.T) {
	tmpDir := t.TempDir()
	mainPath := filepath.Join(tmpDir, "main.asp")
	includePath := filepath.Join(tmpDir, "bad.inc")

	includeSource := `<%
Dim a
a =
%>`
	if err := os.WriteFile(includePath, []byte(includeSource), 0o600); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	mainSource := `<!--#include file="bad.inc"-->
<%
Response.Write "ok"
%>`
	if err := os.WriteFile(mainPath, []byte(mainSource), 0o600); err != nil {
		t.Fatalf("write main failed: %v", err)
	}

	compiler := NewASPCompiler(mainSource)
	compiler.SetSourceName(mainPath)
	err := compiler.Compile()
	if err == nil {
		t.Fatal("expected compile error")
	}

	var syntaxErr *vbscript.VBSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected VBSyntaxError, got %T: %v", err, err)
	}

	if !strings.EqualFold(filepath.Clean(syntaxErr.File), filepath.Clean(includePath)) {
		t.Fatalf("expected mapped file %q, got %q", includePath, syntaxErr.File)
	}

	if syntaxErr.Line != 4 {
		t.Fatalf("expected mapped line 4, got %d", syntaxErr.Line)
	}
}

// TestASPMetadataTypeLibADODBConstants verifies that <!-- METADATA TYPE="TypeLib" --> with the
// ADODB 2.5 UUID makes all ad-prefixed constants available with their correct numeric values,
// and that the directive itself is stripped from the rendered HTML output.
func TestASPMetadataTypeLibADODBConstants(t *testing.T) {
	source := `<html>
<!--
   METADATA
   TYPE="TypeLib"
   NAME="Microsoft ActiveX Data Objects 2.5 Library"
   UUID="{00000205-0000-0010-8000-00AA006D2EA4}"
   VERSION="2.5"
-->
<body><%
Response.Write adInteger & "|" & adVarChar & "|" & adDecimal & "|"
Response.Write adOpenKeyset & "|" & adUseClient & "|" & adLockOptimistic & "|"
Response.Write adFldKeyColumn & "|" & adFldMayBeNull
%></body></html>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	// Verify the METADATA comment was stripped from HTML output.
	if strings.Contains(output.String(), "METADATA") || strings.Contains(output.String(), "TypeLib") {
		t.Fatalf("METADATA directive leaked into HTML output: %q", output.String())
	}

	// Verify all constants have the correct official ADODB 2.5 values.
	const expected = "3|200|14|1|3|3|32768|64"
	if !strings.Contains(output.String(), expected) {
		t.Fatalf("unexpected ADODB constants output:\nexpected to contain: %q\nfull output: %q", expected, output.String())
	}
}

// TestASPTypeLibConstantsAreNotGlobalByDefault verifies non-VBScript typelibrary constants are
// not always-on globals when no METADATA directive is declared.
func TestASPTypeLibConstantsAreNotGlobalByDefault(t *testing.T) {
	source := `<%
Response.Write "x" & CDRom & "|" & Alias
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "x|" {
		t.Fatalf("typelib constants should not be preloaded without METADATA, got output: %q", output.String())
	}
}

// TestASPMetadataTypeLibFSOConstants verifies FSO constants are injected only when
// Microsoft Scripting Runtime METADATA is present.
func TestASPMetadataTypeLibFSOConstants(t *testing.T) {
	source := `<%
Response.Write "before|"
%><!--
METADATA
TYPE="TypeLib"
NAME="Microsoft Scripting Runtime"
UUID="{420B2830-E718-11CF-893D-00A0C9054228}"
VERSION="1.0"
--><%
Response.Write CDRom & "|" & Alias & "|" & ForReading
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if strings.Contains(output.String(), "METADATA") || strings.Contains(output.String(), "TypeLib") {
		t.Fatalf("METADATA directive leaked into HTML output: %q", output.String())
	}

	if !strings.Contains(output.String(), "before|4|64|1") {
		t.Fatalf("unexpected FSO constants output: %q", output.String())
	}
}

// TestASPMetadataTypeLibCDOConstants verifies CDO constants are injected only when
// Collaboration Data Objects METADATA is present.
func TestASPMetadataTypeLibCDOConstants(t *testing.T) {
	source := `<%
Response.Write "x|"
%><!--
METADATA
TYPE="TypeLib"
NAME="Microsoft CDO for Windows 2000 Library"
UUID="{CD000000-8B95-11D1-82DB-00C04FB1625D}"
VERSION="1.21"
--><%
Response.Write cdoSendUsingPort & "|" & cdoBasic & "|" & cdoEncodingBase64
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if strings.Contains(output.String(), "METADATA") || strings.Contains(output.String(), "TypeLib") {
		t.Fatalf("METADATA directive leaked into HTML output: %q", output.String())
	}

	if !strings.Contains(output.String(), "x|2|1|0") {
		t.Fatalf("unexpected CDO constants output: %q", output.String())
	}
}

// TestASPFSOReadAllWithoutParentheses verifies TextStream.ReadAll works when accessed as
// member-get in expression context and remains compatible with Replace(..., vbCrLf, ...).
func TestASPFSOReadAllWithoutParentheses(t *testing.T) {
	rootDir := t.TempDir()
	workspacePath := filepath.Join(rootDir, "workspace")

	source := `<%
Option Explicit
Dim fso, testFile, fileStream, fileContent

Set fso = Server.CreateObject("Scripting.FileSystemObject")
fso.CreateFolder "workspace"

testFile = "workspace\\sample.txt"
Set fileStream = fso.CreateTextFile(testFile, True)
fileStream.WriteLine "Line A"
fileStream.WriteLine "Line B"
fileStream.Close

Set fileStream = fso.OpenTextFile(testFile, 1)
fileContent = fileStream.ReadAll
fileStream.Close

Response.Write Replace(fileContent, vbCrLf, "<br>")
%>`

	compiler := NewASPCompiler(source)
	compiler.SetSourceName(filepath.Join(rootDir, "test.asp"))
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/test.asp")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	got := output.String()
	if !strings.Contains(got, "Line A<br>") || !strings.Contains(got, "Line B") {
		t.Fatalf("unexpected ReadAll/Replace output: %q", got)
	}

	if _, err := os.Stat(filepath.Join(workspacePath, "sample.txt")); err != nil {
		t.Fatalf("expected sample file to exist: %v", err)
	}
}

// TestASPFSOReadLineWithoutParentheses verifies TextStream.ReadLine works when accessed
// without parentheses in expression context (Classic ASP compatibility).
func TestASPFSOReadLineWithoutParentheses(t *testing.T) {
	rootDir := t.TempDir()

	source := `<%
Option Explicit
Dim fso, testFile, fileStream

Set fso = Server.CreateObject("Scripting.FileSystemObject")
fso.CreateFolder "workspace"

testFile = "workspace\\sample.txt"
Set fileStream = fso.CreateTextFile(testFile, True)
fileStream.WriteLine "First"
fileStream.WriteLine "Second"
fileStream.Close

Set fileStream = fso.OpenTextFile(testFile, 1)
Response.Write fileStream.ReadLine & "|" & fileStream.ReadLine
fileStream.Close
%>`

	compiler := NewASPCompiler(source)
	compiler.SetSourceName(filepath.Join(rootDir, "test.asp"))
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/test.asp")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if strings.TrimSpace(output.String()) != "First|Second" {
		t.Fatalf("unexpected ReadLine output: %q", output.String())
	}
}

// TestASPPlusOperatorPageStringRegression verifies page-style string assembly with
// line continuation and '+' renders text rather than numeric zero coercions.
func TestASPPlusOperatorPageStringRegression(t *testing.T) {
	source := `<%
Dim welcomeMessage, apiEndpoint, authToken, requestMethod, fullUrl, authHeader

welcomeMessage = "Welcome to the enhanced VBScript parser! " + _
                 "This version supports colon separators, " + _
                 "plus operators for concatenation, " + _
                 "and line continuation with underscore."
Response.Write welcomeMessage & "|"

apiEndpoint = "https://api.example.com" : authToken = "Bearer ABC123"
requestMethod = "POST"
fullUrl = apiEndpoint + _
          "/v1/users/" + _
          "create?method=" + _
          requestMethod
authHeader = "Authorization: " + authToken

Response.Write fullUrl & "|" & authHeader
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	got := output.String()
	if strings.Contains(got, "0|0") || strings.Contains(got, "Auth: 0") {
		t.Fatalf("unexpected numeric-zero coercion in page-style '+' output: %q", got)
	}
	if !strings.Contains(got, "Welcome to the enhanced VBScript parser!") {
		t.Fatalf("missing welcome message in output: %q", got)
	}
	if !strings.Contains(got, "https://api.example.com/v1/users/create?method=POST") {
		t.Fatalf("missing URL assembly in output: %q", got)
	}
	if !strings.Contains(got, "Authorization: Bearer ABC123") {
		t.Fatalf("missing auth header in output: %q", got)
	}
}

// TestASPInlineCommentBeforeBlockEndPreservesPercentClose verifies Classic ASP
// apostrophe comments do not swallow the closing %> delimiter on the same line,
// and one immediate HTML-leading newline after %> is suppressed.
func TestASPInlineCommentBeforeBlockEndPreservesPercentClose(t *testing.T) {
	source := `<%
Response.Write "before|"
Response.Flush() ' flush buffer mid-page %>
<p>after</p>`

	compiler := NewASPCompiler(source)
	compiler.SetSourceName("/tests/flush_comment.asp")
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "before|<p>after</p>"
	if output.String() != expected {
		t.Fatalf("unexpected output:\nexpected: %q\nactual:   %q", expected, output.String())
	}
}

// TestASPScriptingDictionaryAdvanced verifies Classic ASP-compatible Dictionary behavior
// for CompareMode, object values, nested dictionaries, For Each keys, and Keys/Items
// access without parentheses.
func TestASPScriptingDictionaryAdvanced(t *testing.T) {
	source := `<%
Class Widget
  Public Name
  Public Function Describe()
    Describe = "Widget(" & Name & ")"
  End Function
End Class

Dim d
Set d = Server.CreateObject("Scripting.Dictionary")
d.CompareMode = 1
d.Add "Answer", 42
d.Add "greeting", "hello"

Dim w
Set w = New Widget
w.Name = "alpha"
d.Add "obj", w

Dim nested
Set nested = Server.CreateObject("Scripting.Dictionary")
nested.CompareMode = 1
nested("x") = "y"
d.Add "nested", nested

Response.Write d("answer") & "|"
Response.Write d("GREETING") & "|"

Dim w2
Set w2 = d("OBJ")
Response.Write w2.Describe & "|"

Dim n2
Set n2 = d("nested")
Response.Write n2("X") & "|"

Dim k, iter
iter = ""
For Each k In d
  iter = iter & k & ";"
Next
Response.Write iter & "|"

Dim keysArr, itemsArr
keysArr = d.Keys
itemsArr = d.Items
Response.Write d.Count & "|" & UBound(keysArr) & "|" & UBound(itemsArr) & "|"
Response.Write keysArr(0) & "|" & TypeName(itemsArr(2)) & "|" & TypeName(itemsArr(3))
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	got := output.String()
	if !strings.Contains(got, "42|hello|Widget(alpha)|y|") {
		t.Fatalf("missing advanced dictionary lookup output: %q", got)
	}
	if !strings.Contains(got, "Answer;") || !strings.Contains(got, "greeting;") || !strings.Contains(got, "obj;") || !strings.Contains(got, "nested;") {
		t.Fatalf("missing For Each dictionary keys output: %q", got)
	}
	if !strings.Contains(got, "|4|3|3|") {
		t.Fatalf("unexpected Count/UBound output: %q", got)
	}
	if !strings.Contains(got, "|Answer|Widget|Dictionary") {
		t.Fatalf("unexpected Keys/Items type output: %q", got)
	}
}

// TestASPWithStatementClassPropertyAssignment verifies that With...End With compiled and executed
// correctly for custom class property assignment and read-back.
func TestASPWithStatementClassPropertyAssignment(t *testing.T) {
	source := `<%
Class SystemCtx
    Public ProjectName
    Public Version
End Class
Dim ctx
Set ctx = New SystemCtx
With ctx
    .ProjectName = "AxonASP"
    .Version = "1.0"
End With
Response.Write ctx.ProjectName & "|" & ctx.Version
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	if got := output.String(); got != "AxonASP|1.0" {
		t.Fatalf("expected %q, got %q", "AxonASP|1.0", got)
	}
}

// TestASPWithStatementExpressionRead verifies that '.Prop' in expression context (right-hand side)
// correctly reads the With-subject property.
func TestASPWithStatementExpressionRead(t *testing.T) {
	source := `<%
Class Info
    Public Name
End Class
Dim o
Set o = New Info
o.Name = "Hello"
Dim result
With o
    result = .Name
End With
Response.Write result
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	if got := output.String(); got != "Hello" {
		t.Fatalf("expected %q, got %q", "Hello", got)
	}
}

// TestASPClassPropertyGetForwardReferenceWithArgs verifies that one Property Get
// can call another Property Get declared later in the same class body.
func TestASPClassPropertyGetForwardReferenceWithArgs(t *testing.T) {
	source := `<%
Class LinkBuilder
	Public Property Get getLink(includeArrow)
		getLink = getClickLink(false)
	End Property

	Public Property Get getClickLink(bo)
		getClickLink = "<a href='default.asp?iId=123'>Title</a>"
	End Property
End Class

Dim p
Set p = New LinkBuilder
Response.Write p.getClickLink(false) & "|" & p.getLink(false)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "<a href='default.asp?iId=123'>Title</a>|<a href='default.asp?iId=123'>Title</a>"
	if got := output.String(); got != expected {
		t.Fatalf("expected %q, got %q", expected, got)
	}
}
