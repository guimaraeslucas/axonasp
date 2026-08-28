package axonvm

import (
	"bytes"
	"strconv"
	"testing"

	"g3pix.com.br/axonasp/v2/vbscript"
)

// TestCollectionBuiltin verifies Server.CreateObject("Collection") allows .Add and direct For Each element iteration.
func TestCollectionBuiltin(t *testing.T) {
	source := `<%
Dim col, x
Set col = Server.CreateObject("Collection")
col.Add "hello"
col.Add "world"
For Each x In col
    Response.Write x & "|"
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

	expected := "hello|world|"
	if output.String() != expected {
		t.Fatalf("unexpected output: got %q want %q", output.String(), expected)
	}
}

// TestCollectionCustomClass verifies a custom VBScript class using [DispId(-4)] and m_Items.[_NewEnum] iterates successfully.
func TestCollectionCustomClass(t *testing.T) {
	source := `<%
Class MyCollection
    Private m_Items
    Private Sub Class_Initialize()
        Set m_Items = Server.CreateObject("Collection")
    End Sub
    Public Sub Add(ByVal item)
        m_Items.Add item
    End Sub
    [DispId(-4)]
    Public Property Get NewEnum()
        Set NewEnum = m_Items.[_NewEnum]
    End Property
End Class

Dim col, x
Set col = New MyCollection
col.Add "foo"
col.Add "bar"

For Each x In col
    Response.Write x & "|"
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

	expected := "foo|bar|"
	if output.String() != expected {
		t.Fatalf("unexpected output: got %q want %q", output.String(), expected)
	}
}

// TestCollectionRegression asserts that For Each loops over Scripting.Dictionary (keys) and ADODB.Recordset (fields) still work.
func TestCollectionRegression(t *testing.T) {
	source := `<%
Dim d, k
Set d = CreateObject("Scripting.Dictionary")
d.Add "k1", "v1"
d.Add "k2", "v2"
For Each k In d
    Response.Write k & ":" & d(k) & "|"
Next

Dim rs, f
Set rs = CreateObject("ADODB.Recordset")
rs.Fields.Append "ID", adInteger, , adFldKeyColumn
rs.Fields.Append "Nome", adVarChar, 50, adFldMayBeNull
rs.Open
rs.AddNew
rs("ID") = 42
rs("Nome") = "Regression"
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

	expected := "k1:v1|k2:v2|ID:42|Nome:Regression|"
	if output.String() != expected {
		t.Fatalf("unexpected output: got %q want %q", output.String(), expected)
	}
}

// TestCollectionFailure asserts graceful failure/Classic ASP identical error when For Each is called on a custom class without NewEnum.
func TestCollectionFailure(t *testing.T) {
	source := `<%
On Error Resume Next
Class SimpleClass
End Class

Dim col, x
Set col = New SimpleClass
For Each x In col
Next
Response.Write Err.Number
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

	expected := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.InvalidProcedureCallOrArgument))
	if output.String() != expected {
		t.Fatalf("unexpected output: got %q want %q", output.String(), expected)
	}
}
