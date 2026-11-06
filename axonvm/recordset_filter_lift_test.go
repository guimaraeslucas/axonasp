//go:build !wasm

package axonvm

import (
	"bytes"
	"strings"
	"testing"
)

// A Recordset filter is a view over the retrieved rows, not a destructive
// operation. Measured on IIS/ADO (Server 2022, fabricated client-side
// recordset, https://github.com/saborrie/iis-oracle run 42205539901):
//
//	opened                ->  count=3 rows=apple;pear;plum;
//	Filter = "name = 'pear'"
//	filtered              ->  count=1 rows=pear;
//	Filter = ""
//	lifted-empty          ->  count=3 rows=apple;pear;plum;
//	Filter = 0            (adFilterNone)
//	lifted-zero           ->  count=3 rows=apple;pear;plum;
//	position-after-apply  ->  AbsolutePosition=1 current=pear
//	position-after-lift   ->  AbsolutePosition=1 current=apple
//
// The engine instead overwrote rs.data with the filtered slice, so lifting
// the filter left only the previously matching rows.
func TestRecordsetFilterLiftRestoresTheFullRowSet(t *testing.T) {
	got := runFilterProbePage(t)

	for _, want := range []string{
		"opened -> count=3 rows=apple;pear;plum;",
		"filtered -> count=1 rows=pear;",
		"lifted-empty -> count=3 rows=apple;pear;plum;",
		"lifted-zero -> count=3 rows=apple;pear;plum;",
		"position-after-apply -> pos=1 current=pear",
		"position-after-lift -> pos=1 current=apple",
	} {
		if !strings.Contains(got, want) {
			t.Errorf("want output containing %q\ngot:\n%s", want, got)
		}
	}
}

func runFilterProbePage(t *testing.T) string {
	t.Helper()
	c := NewASPCompiler(`<%
Function Names(rs)
	Dim parts
	parts = ""
	If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
	Do While Not rs.EOF
		parts = parts & rs("name") & ";"
		rs.MoveNext
	Loop
	Names = parts
End Function

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Fields.Append "name", 200, 50
rs.LockType = 3
rs.Open
For Each n In Array("apple", "pear", "plum")
	rs.AddNew
	rs("name") = n
	rs.Update
Next

Response.Write "opened -> count=" & rs.RecordCount & " rows=" & Names(rs) & vbLf
rs.Filter = "name = 'pear'"
Response.Write "filtered -> count=" & rs.RecordCount & " rows=" & Names(rs) & vbLf
rs.Filter = ""
Response.Write "lifted-empty -> count=" & rs.RecordCount & " rows=" & Names(rs) & vbLf
rs.Filter = "name = 'plum'"
rs.Filter = 0
Response.Write "lifted-zero -> count=" & rs.RecordCount & " rows=" & Names(rs) & vbLf
rs.Filter = "name = 'pear'"
Response.Write "position-after-apply -> pos=" & rs.AbsolutePosition & " current=" & rs("name") & vbLf
rs.Filter = ""
Response.Write "position-after-lift -> pos=" & rs.AbsolutePosition & " current=" & rs("name") & vbLf
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
