package axonvm

import (
	"bytes"
	"errors"
	"testing"

	"g3pix.com.br/axonasp/v2/vbscript"
)

// expectedEndHRESULT is 0x800A03F6 expressed as the signed 32-bit number the ASPError
// object reports for VBScript catalog error 1014 (missing End).
const expectedEndHRESULT = -2146827274

// ifBlockResolutionCase describes one block-resolution scenario for If/ElseIf when an
// ASP tag boundary terminates a branch body.
type ifBlockResolutionCase struct {
	name string
	// source is compiled as an ASP page; HTML outside the tags is written verbatim.
	source string
	// wantError is the VBScript catalog code the compilation must fail with.
	// Zero means the source must compile.
	wantError vbscript.VBSyntaxErrorCode
	// wantDescription is the compiler message required when wantError is non-zero.
	wantDescription string
	// wantOutput is the response body required when the source compiles.
	wantOutput string
}

// ifBlockResolutionCases covers single-line If, branch chains (ElseIf/Else) and block If
// heads across ASP tag boundaries. The shapes that must stay distinguishable are:
//
//	single-line If completed by "%>"  -> a later "<% End If %>" is orphaned
//	same-line chain completed by "%>" -> complete like the single-line If; the same
//	                                     "<% End If %>" is orphaned
//	block If head                     -> "<% End If %>" is mandatory
//	inline If nested in a block If    -> the nested statement never claims the outer
//	                                     End If, Else or ElseIf
var ifBlockResolutionCases = []ifBlockResolutionCase{
	{
		name:            "single line if completed by tag boundary rejects following end if",
		source:          `<% Dim c : c = True : If c Then Response.Write "body" %>TAIL<% End If %>`,
		wantError:       vbscript.ExpectedEnd,
		wantDescription: "Expected keyword End",
	},
	{
		name:            "single line if false branch completed by tag boundary rejects following end if",
		source:          `<% Dim c : c = False : If c Then Response.Write "body" %>TAIL<% End If %>`,
		wantError:       vbscript.ExpectedEnd,
		wantDescription: "Expected keyword End",
	},
	{
		name:       "single line if without end if keeps executing following html",
		source:     `<% Dim c : c = True : If c Then Response.Write "body" %>TAIL`,
		wantOutput: "bodyTAIL",
	},
	{
		name:       "single line if false branch without end if skips branch and keeps html",
		source:     `<% Dim c : c = False : If c Then Response.Write "body" %>TAIL`,
		wantOutput: "TAIL",
	},
	{
		// Measured: IIS rejects this construct outright with 800A0401 "Expected end of
		// statement" at the ElseIf — there is no single-line ElseIf in VBScript
		// (https://github.com/saborrie/iis-oracle/actions/runs/42512373768). The engine
		// keeps single-line ElseIf as a superset, so what it reports is the orphan
		// End If: the chain is complete at the tag boundary and nothing after the
		// boundary can attach to it.
		name:            "same line elseif chain completed by tag boundary rejects following end if",
		source:          `<% Dim c : c = 2 : If c = 1 Then Response.Write "one" ElseIf c = 2 Then Response.Write "two" %><% End If %>`,
		wantError:       vbscript.ExpectedEnd,
		wantDescription: "Expected keyword End",
	},
	{
		name:       "block if with inline elseif branch closed by tag boundary executes elseif branch",
		source:     `<% Dim c : c = 2 : If c = 1 Then %><% ElseIf c = 2 Then Response.Write "two" %><% End If %>`,
		wantOutput: "two",
	},
	{
		// Measured: IIS rejects the End If with 800A0400 "Expected statement" — the
		// same-line If/Else chain is complete at the tag boundary, exactly like the
		// plain single-line If (https://github.com/saborrie/iis-oracle/actions/runs/42512373768).
		name:            "same line else chain completed by tag boundary rejects following end if",
		source:          `<% Dim c : c = 0 : If c = 1 Then Response.Write "yes" Else Response.Write "no" %><% End If %>`,
		wantError:       vbscript.ExpectedEnd,
		wantDescription: "Expected keyword End",
	},
	{
		// Measured: IIS rejects the Else with 800A03EA — "If c Then <stmt>" is complete
		// at the tag boundary and a following Else never attaches to it, whether or not
		// the boundary is empty (https://github.com/saborrie/iis-oracle/actions/runs/42512101127,
		// https://github.com/saborrie/iis-oracle/actions/runs/42512053123). With an
		// enclosing block If open the Else binds there instead; see
		// inline_if_else_binding_test.go.
		name:            "detached else after a completed single line if is a syntax error",
		source:          `<% Dim c : c = 0 : If c = 1 Then Response.Write "yes" %><% Else %><% Response.Write "no" %><% End If %>`,
		wantError:       vbscript.SyntaxError,
		wantDescription: "Syntax error: Unexpected token *vbscript.KeywordToken in expression",
	},
	{
		name:       "block if with inline elseif branch closed on same line executes elseif branch",
		source:     `<% Dim c : c = 2 : If c = 1 Then %><% ElseIf c = 2 Then Response.Write "two" : End If %>`,
		wantOutput: "two",
	},
	{
		name:       "inline if without end if keeps elseif and else chain valid",
		source:     `<% Dim n : n = 3 : If n = 1 Then Response.Write "one" ElseIf n = 3 Then Response.Write "three" Else Response.Write "other" %>`,
		wantOutput: "three",
	},
	{
		name:       "same line colon end if still closes inline if",
		source:     `<% Dim s : s = "abc?x=1" : If InStr(s, "?") > 0 Then s = Left(s, InStr(s, "?") - 1) : End If : Response.Write s %>`,
		wantOutput: "abc",
	},
	{
		name:       "single line if nested in block if inner true executes inner branch",
		source:     `<% Dim t : t = "" : If True Then %><% If True Then t = t & "I" %><% t = t & "B" %><% End If %><% Response.Write t %>`,
		wantOutput: "IB",
	},
	{
		name:       "single line if nested in block if inner false skips inner branch",
		source:     `<% Dim t : t = "" : If True Then %><% If False Then t = t & "I" %><% t = t & "B" %><% End If %><% Response.Write t %>`,
		wantOutput: "B",
	},
	{
		name:       "same line else branch ends at line end",
		source:     "<% Dim a : a = 1\nIf a = 1 Then a = 2 Else a = 3\nResponse.Write \"after\" %>",
		wantOutput: "after",
	},
	{
		name:       "same line elseif branch ends at line end",
		source:     "<% Dim a : a = 2\nIf a = 1 Then a = 5 ElseIf a = 3 Then a = 6\nResponse.Write a %>",
		wantOutput: "2",
	},
	{
		name:       "same line else in block if does not claim outer else",
		source:     "<% Dim x, a : x = 1 : a = 1\nIf x = 1 Then\nIf a = 1 Then a = 2 Else a = 3\nResponse.Write \"in\"\nElse\nResponse.Write \"out\"\nEnd If %>",
		wantOutput: "in",
	},
	{
		name:       "same line else ending at tag boundary does not claim outer else",
		source:     "<% Dim x, a : x = 1 : a = 1 %><% If x = 1 Then %>[<% If a = 1 Then Response.Write \"y\" Else Response.Write \"n\" %>]<% Else %>out<% End If %>",
		wantOutput: "[y]",
	},
	{
		name: "inline if closer does not consume end sub",
		source: `<%
Dim r
Sub SetR()
  If True Then r = "ok"
End Sub
SetR
Response.Write r
%>`,
		wantOutput: "ok",
	},
	{
		name: "inline if closer does not consume end function",
		source: `<%
Function Ext(str)
  Dim pos : pos = InStr(str, ".")
  If pos > 0 Then Ext = Mid(str, pos + 1)
End Function
Response.Write Ext("document.pdf")
%>`,
		wantOutput: "pdf",
	},
}

// requireCompileError compiles one ASP source and asserts that it fails with the given
// VBScript catalog code, with the ASPError number derived from the same matrix.
func requireCompileError(t *testing.T, source string, wantCode vbscript.VBSyntaxErrorCode) *vbscript.VBSyntaxError {
	t.Helper()

	compiler := NewASPCompiler(source)
	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compilation error %d (%s), got success for %q", wantCode, wantCode, source)
	}

	var vbErr *vbscript.VBSyntaxError
	if !errors.As(err, &vbErr) {
		t.Fatalf("expected *vbscript.VBSyntaxError, got %T: %v", err, err)
	}
	if vbErr.Code != wantCode {
		t.Fatalf("expected error code %d (%s), got %d (%s)", wantCode, wantCode, vbErr.Code, vbErr.Code)
	}
	if wantNumber := vbscript.HRESULTFromVBScriptCode(wantCode); vbErr.Number != wantNumber {
		t.Fatalf("expected error number %d (%s), got %d", wantNumber, vbscript.HRESULTHexFromVBScriptCode(wantCode), vbErr.Number)
	}
	return vbErr
}

// requireOrphanEndIfError compiles one ASP source and asserts that it fails with the
// VBScript orphan End If error: catalog 1014, ASPError 0x800A03F6.
func requireOrphanEndIfError(t *testing.T, source string) *vbscript.VBSyntaxError {
	t.Helper()

	vbErr := requireCompileError(t, source, vbscript.ExpectedEnd)
	if vbErr.Description != "Expected keyword End" {
		t.Fatalf("expected description %q, got %q", "Expected keyword End", vbErr.Description)
	}
	return vbErr
}

// TestExpectedEndErrorCodeMatchesMatrix locks the HRESULT reported for the orphan End If
// error so the ASPError number stays the documented 0x800A03F6.
func TestExpectedEndErrorCodeMatchesMatrix(t *testing.T) {
	if got := vbscript.HRESULTFromVBScriptCode(vbscript.ExpectedEnd); got != expectedEndHRESULT {
		t.Fatalf("ExpectedEnd HRESULT changed: got %d (0x%X), want %d (0x800A03F6)", got, got, expectedEndHRESULT)
	}
}

// TestIfBlockResolutionAcrossTagBoundaries compiles and runs every If block-resolution
// scenario, asserting either the exact compilation error or the exact response body.
func TestIfBlockResolutionAcrossTagBoundaries(t *testing.T) {
	for _, tc := range ifBlockResolutionCases {
		t.Run(tc.name, func(t *testing.T) {
			if tc.wantError != 0 {
				vbErr := requireCompileError(t, tc.source, tc.wantError)
				if vbErr.Description != tc.wantDescription {
					t.Fatalf("expected description %q, got %q", tc.wantDescription, vbErr.Description)
				}
				return
			}

			compiler := NewASPCompiler(tc.source)
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

			if output.String() != tc.wantOutput {
				t.Fatalf("unexpected output: got %q want %q", output.String(), tc.wantOutput)
			}
		})
	}
}
