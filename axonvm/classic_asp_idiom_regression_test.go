//go:build !wasm

package axonvm

import (
	"bytes"
	"fmt"
	"path/filepath"
	"strings"
	"testing"
)

// End-to-end regression tests for classic ASP idioms, distilled from a
// production codebase of ~580 pages running on AxonASP. Each test is a
// synthetic page exercising one idiom family the way real pages combine it:
// compile, run against a real SQLite database, compare rendered output.
//
// The point is to pin whole-pipeline behaviour, not units: most of these
// idioms appear in dozens to hundreds of pages each, so a regression in any
// of them is a mass breakage that no single-function test catches.

// runIdiomPage compiles and runs one ASP page and returns what it rendered.
func runIdiomPage(t *testing.T, source string) string {
	t.Helper()
	c := NewASPCompiler(source)
	if err := c.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVMFromCompiler(c)
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("run failed: %v", err)
	}
	host.Response().Flush()
	return out.String()
}

// seededPage wraps a page body with a connection to a fresh SQLite database
// holding one small, known table.
func seededPage(t *testing.T, body string) string {
	t.Helper()
	dbPath := filepath.Join(t.TempDir(), "idioms.db")
	return fmt.Sprintf(`<%%
set conn = Server.CreateObject("ADODB.Connection")
conn.Open "sqlite:%s"
conn.Execute "CREATE TABLE fruit (id INTEGER PRIMARY KEY, name TEXT NOT NULL, price REAL, note TEXT)"
conn.Execute "INSERT INTO fruit (name, price, note) VALUES ('apple', 1.5, 'crisp'), ('pear', 2, NULL), ('plum', 3.25, 'ripe')"
%s
conn.Close
set conn = Nothing
%%>`, dbPath, body)
}

func assertRendered(t *testing.T, got, want string) {
	t.Helper()
	if !strings.Contains(got, want) {
		t.Errorf("want output containing %q, got %q", want, got)
	}
}

// The workhorse read loop: Execute, While Not EOF ... MoveNext ... Wend,
// Close, Set Nothing.
func TestIdiomWhileNotEOFReadLoop(t *testing.T) {
	got := runIdiomPage(t, seededPage(t, `
set rs = conn.Execute("SELECT name FROM fruit ORDER BY id")
While Not rs.EOF
	Response.Write rs("name") & ";"
	rs.MoveNext
Wend
rs.Close
set rs = Nothing
Response.Write "done"`))
	assertRendered(t, got, "apple;pear;plum;done")
}

// RecordCount after opening, then MoveFirst to re-read after an iteration.
func TestIdiomRecordCountAndMoveFirst(t *testing.T) {
	got := runIdiomPage(t, seededPage(t, `
set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT name FROM fruit ORDER BY id", conn
Response.Write "count=" & rs.RecordCount & " "
While Not rs.EOF
	rs.MoveNext
Wend
rs.MoveFirst
Response.Write "first=" & rs("name")
rs.Close`))
	assertRendered(t, got, "count=3 first=apple")
}

// GetRows: fields are the first dimension, rows the second.
func TestIdiomGetRowsMatrix(t *testing.T) {
	got := runIdiomPage(t, seededPage(t, `
set rs = conn.Execute("SELECT id, name FROM fruit ORDER BY id")
arr = rs.GetRows()
Response.Write "fields=" & UBound(arr, 1) & " rows=" & UBound(arr, 2) & " "
Response.Write "cell=" & arr(1, 2)`))
	assertRendered(t, got, "fields=1 rows=2 cell=plum")
}

// BeginTrans/RollbackTrans discards the insert; CommitTrans keeps it.
func TestIdiomTransactionRollbackAndCommit(t *testing.T) {
	got := runIdiomPage(t, seededPage(t, `
conn.BeginTrans
conn.Execute "INSERT INTO fruit (name) VALUES ('mango')"
conn.RollbackTrans
set rs = conn.Execute("SELECT COUNT(*) AS n FROM fruit")
Response.Write "afterRollback=" & rs("n") & " "
conn.BeginTrans
conn.Execute "INSERT INTO fruit (name) VALUES ('mango')"
conn.CommitTrans
set rs = conn.Execute("SELECT COUNT(*) AS n FROM fruit")
Response.Write "afterCommit=" & rs("n")`))
	assertRendered(t, got, "afterRollback=3 afterCommit=4")
}

// On Error Resume Next around a failing statement: the error is trapped,
// Err carries it, Err.Clear resets it, and the page carries on.
func TestIdiomOnErrorResumeNextTrapsAQueryError(t *testing.T) {
	got := runIdiomPage(t, seededPage(t, `
On Error Resume Next
conn.Execute "SELECT * FROM no_such_table"
If Err.Number <> 0 Then
	Response.Write "trapped "
End If
Err.Clear
Response.Write "cleared=" & Err.Number & " "
On Error Goto 0
Response.Write "alive"`))
	assertRendered(t, got, "trapped cleared=0 alive")
}

// The string-function battery that renders most list pages.
func TestIdiomStringFunctions(t *testing.T) {
	got := runIdiomPage(t, `<%
parts = Split("alpha,beta,gamma", ",")
Response.Write "n=" & UBound(parts) & " "
Response.Write "joined=" & Join(parts, "-") & " "
Response.Write "replaced=" & Replace("aabb", "b", "c") & " "
Response.Write "at=" & InStr("warehouse", "house") & " "
Response.Write "trimmed=[" & Trim("  pad  ") & "] "
Response.Write "mid=" & Mid("abcdef", 3, 2) & " "
Response.Write "sum=" & (CInt("41") + 1)
%>`)
	assertRendered(t, got,
		"n=2 joined=alpha-beta-gamma replaced=aacc at=5 trimmed=[pad] mid=cd sum=42")
}

// Deterministic date arithmetic (no Now()): DateSerial, DateAdd, DateDiff
// and the part extractors.
func TestIdiomDateArithmetic(t *testing.T) {
	got := runIdiomPage(t, `<%
d = DateSerial(2026, 1, 1)
Response.Write "diff=" & DateDiff("d", d, DateSerial(2026, 1, 31)) & " "
Response.Write "day=" & Day(DateAdd("d", 5, d)) & " "
Response.Write "month=" & Month(DateAdd("m", 2, d)) & " "
Response.Write "year=" & Year(DateAdd("yyyy", 1, d))
%>`)
	assertRendered(t, got, "diff=30 day=6 month=3 year=2027")
}
