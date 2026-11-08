//go:build !wasm

package axonvm

import (
	"bytes"
	"strings"
	"testing"
)

// How VBScript treats Null, measured on real IIS rather than derived from
// documentation (https://github.com/saborrie/iis-oracle, cited in
// https://github.com/guimaraeslucas/axonasp/issues/188):
//
//   - & with exactly one Null operand coerces it to a zero-length string and
//     the result is a String; Null & Null stays Null. + propagates Null.
//   - String coercion of Null renders nothing: Response.Write Null writes
//     "" and raises nothing.
//   - The explicit casts (CStr, CBool, CInt, CLng, CDbl, CDate) raise 94.
//   - The string functions propagate Null: Len, LenB, Trim, LTrim, RTrim,
//     UCase, LCase, Left, Right, Mid, InStr, StrComp all return Null.
//     Replace is the measured exception: 94 for Null in any string argument.
//   - If Null Then takes the false branch, without raising.
//
// The engine instead stringified Null as the word "Null", so every page that
// concatenates a nullable database column into its output rendered "Null".
func TestNullSemanticsMatchIIS(t *testing.T) {
	got := runNullMatrixPage(t)

	for _, want := range []string{
		`"a" & Null -> [a] (vartype=8)`,
		`Null & "a" -> [a] (vartype=8)`,
		`Null & Null -> Null (vartype=1)`,
		`Null & Empty -> [] (vartype=8)`,
		`Empty & Null -> [] (vartype=8)`,
		`1 & Null -> [1] (vartype=8)`,
		`Null & 1 -> [1] (vartype=8)`,
		`"a" + Null -> Null (vartype=1)`,
		`1 + Null -> Null (vartype=1)`,
		`Null + Null -> Null (vartype=1)`,
		`Null = Null -> Null (vartype=1)`,
		`IsNull("a" & Null) -> [False] (vartype=11)`,
		`CStr(Null) -> ERR (Invalid use of Null)`,
		`CBool(Null) -> ERR (Invalid use of Null)`,
		`CInt(Null) -> ERR (Invalid use of Null)`,
		`CLng(Null) -> ERR (Invalid use of Null)`,
		`CDbl(Null) -> ERR (Invalid use of Null)`,
		`CDate(Null) -> ERR (Invalid use of Null)`,
		`Len(Null) -> Null (vartype=1)`,
		`LenB(Null) -> Null (vartype=1)`,
		`Trim(Null) -> Null (vartype=1)`,
		`LTrim(Null) -> Null (vartype=1)`,
		`RTrim(Null) -> Null (vartype=1)`,
		`UCase(Null) -> Null (vartype=1)`,
		`LCase(Null) -> Null (vartype=1)`,
		`Left(Null, 1) -> Null (vartype=1)`,
		`Right(Null, 1) -> Null (vartype=1)`,
		`Mid(Null, 1) -> Null (vartype=1)`,
		`InStr(Null, "a") -> Null (vartype=1)`,
		`InStr("a", Null) -> Null (vartype=1)`,
		`InStr(1, Null, "a") -> Null (vartype=1)`,
		`StrComp(Null, "a") -> Null (vartype=1)`,
		`Replace(Null, "a", "b") -> ERR (Invalid use of Null)`,
		`Replace("aa", Null, "b") -> ERR (Invalid use of Null)`,
		`Replace("aa", "a", Null) -> ERR (Invalid use of Null)`,
		`"a" & Null & "b" -> [ab] (vartype=8)`,
		`Len("" & Null) -> [0]`,
		"If Null Then -> false-branch",
		"writeN-concat -> []",
		"direct-write -> [] err=0",
	} {
		if !strings.Contains(got, want) {
			t.Errorf("want output containing %q", want)
		}
	}
	if t.Failed() {
		t.Logf("full output:\n%s", got)
	}
}

func runNullMatrixPage(t *testing.T) string {
	t.Helper()
	c := NewASPCompiler(`<%
Function Describe(expr)
	Dim v
	On Error Resume Next
	Err.Clear
	v = Eval(expr)
	If Err.Number <> 0 Then
		Describe = "ERR (" & Err.Description & ")"
		Err.Clear
	ElseIf IsNull(v) Then
		Describe = "Null (vartype=" & VarType(v) & ")"
	ElseIf IsEmpty(v) Then
		Describe = "Empty (vartype=" & VarType(v) & ")"
	Else
		Describe = "[" & CStr(v) & "] (vartype=" & VarType(v) & ")"
	End If
	On Error Goto 0
End Function
cases = Array( _
	"""a"" & Null", "Null & ""a""", "Null & Null", "Null & Empty", "Empty & Null", _
	"1 & Null", "Null & 1", """a"" + Null", "1 + Null", "Null + Null", _
	"Null = Null", "IsNull(""a"" & Null)", _
	"CStr(Null)", "CBool(Null)", "CInt(Null)", "CLng(Null)", "CDbl(Null)", "CDate(Null)", _
	"Len(Null)", "LenB(Null)", "Trim(Null)", "LTrim(Null)", "RTrim(Null)", _
	"UCase(Null)", "LCase(Null)", "Left(Null, 1)", "Right(Null, 1)", "Mid(Null, 1)", _
	"InStr(Null, ""a"")", "InStr(""a"", Null)", "InStr(1, Null, ""a"")", _
	"StrComp(Null, ""a"")", _
	"Replace(Null, ""a"", ""b"")", "Replace(""aa"", Null, ""b"")", "Replace(""aa"", ""a"", Null)", _
	"""a"" & Null & ""b""", "Len("""" & Null)" _
)
For Each c In cases
	Response.Write c & " -> " & Describe(c) & vbLf
Next
On Error Resume Next
Err.Clear
hit = "no-branch-ran"
If Null Then hit = "true-branch" Else hit = "false-branch"
If Err.Number <> 0 Then
	Response.Write "If Null Then -> ERR; " & hit & vbLf
Else
	Response.Write "If Null Then -> " & hit & vbLf
End If
x = Null
Response.Write "writeN-concat -> [" & x & "]" & vbLf
Err.Clear
Response.Write "direct-write -> ["
Response.Write Null
Response.Write "] err=" & Err.Number & vbLf
%>`)
	if err := c.Compile(); err != nil {
		t.Fatalf("compile: %v", err)
	}
	vm := NewVMFromCompiler(c)
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("run: %v", err)
	}
	host.Response().Flush()
	return out.String()
}
