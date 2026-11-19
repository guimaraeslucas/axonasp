//go:build !wasm && !lib_adodb_disabled

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

import (
	"bytes"
	"strconv"
	"strings"
	"testing"

	"g3pix.com.br/axonasp/v2/vbscript"
)

// runADODBFieldScript compiles and executes an ASP snippet, returning its output, VM, and any fatal runtime error.
func runADODBFieldScript(t *testing.T, source string) (string, *VM, error) {
	t.Helper()
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	err := vm.Run()
	host.Response().Flush()
	return output.String(), vm, err
}

// TestADODBFieldMinimalReproductionParity verifies rs("descrizione").Value and
// rs("descrizione") both return the underlying value with Err.Number = 0.
func TestADODBFieldMinimalReproductionParity(t *testing.T) {
	source := `<%
On Error Resume Next

Dim conn, rs, fieldValue
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "sqlite::memory:"
Set rs = conn.Execute( _
    "SELECT CAST('Murano Glass Lamp' AS nvarchar(100)) AS descrizione")

Err.Clear
fieldValue = rs("descrizione").Value

Response.Write "Explicit .Value: " & fieldValue & "|"
Response.Write "Error: " & Err.Number & " - " & Err.Description & "|"

Err.Clear
fieldValue = rs("descrizione")

Response.Write "Direct access: " & fieldValue & "|"
Response.Write "Error: " & Err.Number

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>`

	got, _, err := runADODBFieldScript(t, source)
	if err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	want := "Explicit .Value: Murano Glass Lamp|Error: 0 - |Direct access: Murano Glass Lamp|Error: 0"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestADODBFieldAccessFormsAndCaseInsensitivity verifies direct access, .Value,
// .value, .VALUE, ordinal access, rs.Fields(...).Value, and rs.Fields.Item(...).Value.
func TestADODBFieldAccessFormsAndCaseInsensitivity(t *testing.T) {
	source := `<%
Dim conn, rs
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "sqlite::memory:"
Set rs = conn.Execute("SELECT CAST('Murano Glass Lamp' AS nvarchar(100)) AS descrizione")

Response.Write rs("descrizione") & "|"
Response.Write rs("descrizione").Value & "|"
Response.Write rs("descrizione").value & "|"
Response.Write rs("DESCRIZIONE").VALUE & "|"
Response.Write rs(0) & "|"
Response.Write rs(0).Value & "|"
Response.Write rs.Fields("descrizione") & "|"
Response.Write rs.Fields("descrizione").Value & "|"
Response.Write rs.Fields(0).value & "|"
Response.Write rs.Fields.Item("descrizione").Value & "|"
Response.Write rs.Fields.Item(0).VALUE

rs.Close
conn.Close
%>`

	got, _, err := runADODBFieldScript(t, source)
	if err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	want := strings.TrimSuffix(strings.Repeat("Murano Glass Lamp|", 11), "|")
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestADODBFieldInExpressionsCallsAndWith verifies rs("col").Value and rs("col")
// work in nested function calls, parenless Sub calls, arithmetic, comparisons, and With blocks.
func TestADODBFieldInExpressionsCallsAndWith(t *testing.T) {
	source := `<%
Function TransformTitle(s)
	TransformTitle = UCase(Trim(s)) & "!"
End Function

Sub EmitLine(v)
	Response.Write v & "|"
End Sub

Dim conn, rs, f
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "sqlite::memory:"
conn.Execute "CREATE TABLE items (descrizione NVARCHAR(100), qty INTEGER, active BOOLEAN)"
conn.Execute "INSERT INTO items VALUES ('  Murano Glass Lamp  ', 7, 1)"
Set rs = conn.Execute("SELECT descrizione, qty, active FROM items")

Response.Write TransformTitle(rs("descrizione").Value) & "|"
Response.Write TransformTitle(rs("descrizione")) & "|"
EmitLine rs("descrizione").Value
EmitLine rs("descrizione")
Response.Write((rs("qty").Value + 5) & "|")
Response.Write((rs("qty") + 5) & "|")
Response.Write((rs("qty").Value = 7) & "|")
Response.Write((rs("qty") = 7) & "|")

With rs("descrizione")
	Dim withLine
	withLine = .Name & "=" & Trim(.Value)
	Response.Write withLine & "|"
End With

Set f = rs("descrizione")
Response.Write IsObject(f) & "|" & TypeName(f) & "|"
Response.Write((rs("descrizione") Is rs.Fields("descrizione")) & "|")
Response.Write((rs(0) Is rs("descrizione")) & "|")
Response.Write Len(f) & "|" & TransformTitle(f)

rs.Close
conn.Close
%>`

	got, _, err := runADODBFieldScript(t, source)
	if err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	want := "MURANO GLASS LAMP!|MURANO GLASS LAMP!|  Murano Glass Lamp  |  Murano Glass Lamp  |12|12|True|True|descrizione=Murano Glass Lamp|True|Field|True|True|21|MURANO GLASS LAMP!"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestADODBFieldVariableCoercionAndCursorTracking verifies that holding a Field proxy
// via Set tracks MoveNext updates and coerces in conditionals, arithmetic, and builtins.
func TestADODBFieldVariableCoercionAndCursorTracking(t *testing.T) {
	source := `<!-- METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" --><%
Dim rs, fName, fCount
Set rs = CreateObject("ADODB.Recordset")
rs.Fields.Append "Name", adVarChar, 50, adFldMayBeNull
rs.Fields.Append "Qty", adInteger, 4, adFldMayBeNull
rs.Open
rs.AddNew
rs("Name") = "First"
rs("Qty") = 10
rs.Update
rs.AddNew
rs("Name") = "Second"
rs("Qty") = 0
rs.Update
rs.MoveFirst

Set fName = rs("Name")
Set fCount = rs("Qty")
If fCount Then Response.Write "truthy:"
Response.Write fName & ":" & (fCount + 2) & ":" & UCase(fName) & "|"

rs.MoveNext
If Not fCount Then Response.Write "falsy:"
If fName = "Second" Then Response.Write "eq:"
Response.Write fName & ":" & (fCount + 2)
rs.Close
%>`

	got, _, err := runADODBFieldScript(t, source)
	if err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	want := "truthy:First:12:FIRST|falsy:eq:Second:2"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestADODBFieldNullAndMultiTypeValues verifies SQL NULL, strings, integers, doubles,
// booleans, dates, and binary blobs work through both direct and explicit .Value paths.
func TestADODBFieldNullAndMultiTypeValues(t *testing.T) {
	source := `<%
Dim conn, rs, fNull
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "sqlite::memory:"
conn.Execute "CREATE TABLE typed_rows (txt NVARCHAR(100), n INTEGER, flt DECIMAL(10,2), b BOOLEAN, dt DATETIME, bin BLOB, nullable_col TEXT)"
conn.Execute "INSERT INTO typed_rows VALUES ('lamp', 42, 19.75, 1, '2026-11-19 10:00:00', X'4142', NULL)"
Set rs = conn.Execute("SELECT txt, n, flt, b, dt, bin, nullable_col FROM typed_rows")

Set fNull = rs("nullable_col")
Response.Write IsNull(rs("nullable_col")) & "|"
Response.Write IsNull(rs("nullable_col").Value) & "|"
Response.Write IsNull(rs.Fields("nullable_col").Value) & "|"
Response.Write IsNull(fNull) & "|"
Response.Write "[" & rs("nullable_col").Value & "]|"
Response.Write "A=" & rs("nullable_col").ActualSize & "|"
Response.Write rs("txt").Value & "|"
Response.Write((rs("n").Value + 8) & "|")
Response.Write rs("flt").Value & "|"
Response.Write CBool(rs("b").Value) & "|"
Response.Write((rs("dt").Value <> "") & "|")
Response.Write (rs("bin").ActualSize > 0)

rs.Close
conn.Close
%>`

	got, _, err := runADODBFieldScript(t, source)
	if err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	want := "True|True|True|True|[]|A=0|lamp|50|19.75|True|True|True"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestADODBFieldMetadataAndWriteForms verifies canonical .Name casing, .Type mapping,
// .DefinedSize, .ActualSize, .NumericScale, and all field assignment forms.
func TestADODBFieldMetadataAndWriteForms(t *testing.T) {
	if got := adodbTypeFromDatabaseType("NVARCHAR(100)"); got != 202 {
		t.Fatalf("adodbTypeFromDatabaseType(NVARCHAR(100)) = %d, want 202", got)
	}
	if got := adodbTypeFromDatabaseType("DECIMAL(10,2)"); got != 131 {
		t.Fatalf("adodbTypeFromDatabaseType(DECIMAL(10,2)) = %d, want 131", got)
	}

	source := `<!-- METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" --><%
Dim rs
Set rs = CreateObject("ADODB.Recordset")
rs.Fields.Append "TitleName", adVarWChar, 100, adFldMayBeNull
rs.Fields.Append "CopyName", adVarWChar, 100, adFldMayBeNull
rs.Fields.Append "Price", adDecimal, 14, adFldMayBeNull
rs.Fields("Price").NumericScale = 2
rs.Open
rs.AddNew
rs("TitleName") = "Alpha"
rs("CopyName").Value = "Beta"
rs.Update

Response.Write rs("titlename").Name & "|"
Response.Write rs("titlename").Type & "|"
Response.Write rs("titlename").DefinedSize & "|"
Response.Write rs("titlename").ActualSize & "|"
Response.Write rs("price").NumericScale & "|"

rs.Fields("TitleName") = "Gamma"
rs.Fields("CopyName").Value = "Delta"
Response.Write rs("TitleName").Value & "|" & rs("CopyName").Value & "|"

rs.Fields.Item("TitleName") = "Epsilon"
rs("CopyName") = rs("TitleName")
Response.Write rs("TitleName").Value & "|" & rs("CopyName").Value
rs.Close
%>`

	got, _, err := runADODBFieldScript(t, source)
	if err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	want := "TitleName|202|100|5|2|Gamma|Delta|Epsilon|Epsilon"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestADODBFieldMissingAndObjectRequiredErrors verifies missing/out-of-range fields raise
// ADO 3265 without secondary Object Required clobbering, while genuine invalid dereferences
// preserve -2146827864 (0x800A01A8 "Object required").
func TestADODBFieldMissingAndObjectRequiredErrors(t *testing.T) {
	itemNotFound := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.ADODBItemNotFound))
	objRequired := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.CouldNotFindTargetObject))

	source := `<%
Function TransformTitle(s)
	Response.Write "CALLED:"
	TransformTitle = UCase(s)
End Function

On Error Resume Next
Dim conn, rs, v, emptyVar, nullObj
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "sqlite::memory:"
Set rs = conn.Execute("SELECT 'Lamp' AS title")

v = "keep"
Err.Clear
v = rs("missing").Value
Response.Write "M1=" & Err.Number & ":" & (InStr(Err.Description, "Item cannot be found") > 0) & ":" & v & "|"

Err.Clear
v = rs("missing")
Response.Write "M2=" & Err.Number & ":" & v & "|"

Err.Clear
v = rs.Fields("missing").Value
Response.Write "M3=" & Err.Number & "|"

Err.Clear
v = rs(99).Value
Response.Write "M4=" & Err.Number & "|"

Err.Clear
v = rs(-1).Value
Response.Write "M5=" & Err.Number & "|"

Err.Clear
v = TransformTitle(rs("missing").Value)
Response.Write "M6=" & Err.Number & ":" & v & "|"

Dim item, loopHits
loopHits = 0
For Each item In Array("a", "b")
	Err.Clear
	v = rs("missing").Value
	loopHits = loopHits + 1
Next
Response.Write "LOOP=" & loopHits & ":" & Err.Number & "|"

Set nullObj = Nothing
Err.Clear
v = nullObj.Value
Response.Write "O1=" & Err.Number & "|"

Err.Clear
v = emptyVar.Value
Response.Write "O2=" & Err.Number & "|"

Err.Clear
v = "scalar".Value
Response.Write "O3=" & Err.Number

rs.Close
conn.Close
%>`

	got, _, err := runADODBFieldScript(t, source)
	if err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	want := "M1=" + itemNotFound + ":True:keep|" +
		"M2=" + itemNotFound + ":keep|" +
		"M3=" + itemNotFound + "|" +
		"M4=" + itemNotFound + "|" +
		"M5=" + itemNotFound + "|" +
		"M6=" + itemNotFound + ":keep|" +
		"LOOP=2:" + itemNotFound + "|" +
		"O1=" + objRequired + "|" +
		"O2=" + objRequired + "|" +
		"O3=" + objRequired
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestADODBFieldUncaughtMissingRaisesItemNotFound verifies missing field access without
// On Error Resume Next propagates vbscript.ADODBItemNotFound (3265).
func TestADODBFieldUncaughtMissingRaisesItemNotFound(t *testing.T) {
	source := `<%
Dim conn, rs
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "sqlite::memory:"
Set rs = conn.Execute("SELECT 'Lamp' AS title")
Response.Write rs("missing").Value
%>`

	_, _, err := runADODBFieldScript(t, source)
	if err == nil {
		t.Fatal("expected uncaught error for missing field, got nil")
	}
	vme, ok := err.(*VMError)
	if !ok {
		t.Fatalf("expected *VMError, got %T (%v)", err, err)
	}
	if vme.Code != vbscript.ADODBItemNotFound {
		t.Fatalf("expected error code %d, got %d (%s)", vbscript.ADODBItemNotFound, vme.Code, vme.Description)
	}
}

// TestJScriptADODBFieldExplicitValueAndOperators verifies JScript rs("field").Value,
// .Name, .Type, string/numeric + addition, loose equality, strict null, and try/catch.
func TestJScriptADODBFieldExplicitValueAndOperators(t *testing.T) {
	source := `<%@ Language="JScript" %><%
var conn = Server.CreateObject("ADODB.Connection");
conn.Open("sqlite::memory:");
conn.Execute("CREATE TABLE products (descrizione NVARCHAR(100), qty INTEGER, notes TEXT)");
conn.Execute("INSERT INTO products VALUES ('Murano Glass Lamp', 7, NULL)");
var rs = conn.Execute("SELECT descrizione, qty, notes FROM products");

var out = [];
out.push(rs("descrizione").Value);
out.push(rs("descrizione").Name);
out.push(String(rs("descrizione").Type));
out.push(String(rs("descrizione")));
out.push(rs("descrizione") + "!");
out.push(rs("qty") + rs("qty"));
out.push(String(rs("descrizione") == "Murano Glass Lamp"));
out.push(String(rs("notes").Value === null));

var caught = false;
try {
    var bad = rs("missing").Value;
} catch (e) {
    caught = true;
}
out.push(String(caught));

rs.Close();
conn.Close();
Response.Write(out.join("|"));
%>`

	got, _, err := runADODBFieldScript(t, source)
	if err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	want := "Murano Glass Lamp|descrizione|202|Murano Glass Lamp|Murano Glass Lamp!|14|true|true|true"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestADODBFieldProxyCachedPerColumn verifies iterating 1,000 rows reuses one
// Field proxy per column instead of allocating one per row cell.
func TestADODBFieldProxyCachedPerColumn(t *testing.T) {
	source := `<!-- METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" --><%
Dim rs, i, totalLen
Set rs = CreateObject("ADODB.Recordset")
rs.Fields.Append "ID", adInteger, 4
rs.Fields.Append "Title", adVarChar, 50
rs.Open
For i = 1 To 1000
	rs.AddNew
	rs("ID") = i
	rs("Title") = "Item " & i
	rs.Update
Next
rs.MoveFirst
totalLen = 0
Do While Not rs.EOF
	totalLen = totalLen + Len(rs("Title").Value) + rs("ID").Value
	rs.MoveNext
Loop
Response.Write totalLen
rs.Close
%>`

	got, vm, err := runADODBFieldScript(t, source)
	if err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	if got == "" || got == "0" {
		t.Fatalf("unexpected totalLen output: %q", got)
	}
	if count := len(vm.adodbFieldItems); count != 2 {
		t.Fatalf("expected exactly 2 cached Field proxies across 1,000 rows, got %d", count)
	}
}
