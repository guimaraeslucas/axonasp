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
//	branch chain completed by "%>"    -> "<% End If %>" closes the open block
//	block If head                     -> "<% End If %>" is mandatory
//	inline If nested in a block If    -> the nested statement never claims the outer End If
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
		name:       "inline elseif chain closed by tag boundary executes elseif branch",
		source:     `<% Dim c : c = 2 : If c = 1 Then Response.Write "one" ElseIf c = 2 Then Response.Write "two" %><% End If %>`,
		wantOutput: "two",
	},
	{
		name:       "block if with inline elseif branch closed by tag boundary executes elseif branch",
		source:     `<% Dim c : c = 2 : If c = 1 Then %><% ElseIf c = 2 Then Response.Write "two" %><% End If %>`,
		wantOutput: "two",
	},
	{
		name:       "inline else chain closed by tag boundary executes else branch",
		source:     `<% Dim c : c = 0 : If c = 1 Then Response.Write "yes" Else Response.Write "no" %><% End If %>`,
		wantOutput: "no",
	},
	{
		name:       "else branch spread across tag boundaries executes else branch",
		source:     `<% Dim c : c = 0 : If c = 1 Then Response.Write "yes" %><% Else %><% Response.Write "no" %><% End If %>`,
		wantOutput: "no",
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

// requireOrphanEndIfError compiles one ASP source and asserts that it fails with the
// VBScript orphan End If error: catalog 1014, ASPError 0x800A03F6.
func requireOrphanEndIfError(t *testing.T, source string) *vbscript.VBSyntaxError {
	t.Helper()

	compiler := NewASPCompiler(source)
	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected orphan End If compilation error, got success for %q", source)
	}

	var vbErr *vbscript.VBSyntaxError
	if !errors.As(err, &vbErr) {
		t.Fatalf("expected *vbscript.VBSyntaxError, got %T: %v", err, err)
	}
	if vbErr.Code != vbscript.ExpectedEnd {
		t.Fatalf("expected error code %d (%s), got %d (%s)", vbscript.ExpectedEnd, vbscript.ExpectedEnd, vbErr.Code, vbErr.Code)
	}
	if vbErr.Number != expectedEndHRESULT {
		t.Fatalf("expected error number %d (0x800A03F6), got %d", expectedEndHRESULT, vbErr.Number)
	}
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
				vbErr := requireOrphanEndIfError(t, tc.source)
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
