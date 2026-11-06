//go:build !wasm

package axonvm

import (
	"path/filepath"
	"testing"
)

// Real ADO executes the recordset's Source property when Open is given no
// source argument:
//
//	rs.Source = "SELECT ..."
//	rs.Open()
//
// The engine only copied a non-empty Open argument into rs.source and never
// fell back the other way, so the no-argument form sent an empty query to the
// driver and came back EOF with no error. Classic ASP code uses this idiom
// pervasively.
func TestRecordsetOpenWithoutArgumentsExecutesTheSourceProperty(t *testing.T) {
	vm, rs := newSourcePropertyTestRecordset(t)

	vm.dispatchMemberSet(rs.Num, "Source", NewString("SELECT name FROM people WHERE id = 2"))
	vm.dispatchNativeCall(rs.Num, "Open", nil)

	if state := vm.dispatchMemberGet(rs, "State"); state.Type != VTInteger || state.Num != adStateOpen {
		t.Fatalf("recordset did not open: State = %#v", state)
	}
	if eof := vm.dispatchMemberGet(rs, "EOF"); eof.Type != VTBool || eof.Num != 0 {
		t.Fatalf("no-argument Open() ran an empty query instead of the Source "+
			"property: EOF = %#v", eof)
	}
	if name := vm.dispatchMemberGet(rs, "name"); name.Type != VTString || name.Str != "second" {
		t.Fatalf("wrong row came back: name = %#v", name)
	}
}

// The already-working form must stay as it is: an explicit source argument
// wins over a previously assigned Source property, and overwrites it.
func TestRecordsetOpenArgumentStillOverridesTheSourceProperty(t *testing.T) {
	vm, rs := newSourcePropertyTestRecordset(t)

	vm.dispatchMemberSet(rs.Num, "Source", NewString("SELECT name FROM people WHERE id = 2"))
	const argued = "SELECT name FROM people WHERE id = 1"
	vm.dispatchNativeCall(rs.Num, "Open", []Value{NewString(argued)})

	if name := vm.dispatchMemberGet(rs, "name"); name.Type != VTString || name.Str != "first" {
		t.Fatalf("the Open argument did not win over the Source property: name = %#v", name)
	}
	if source := vm.dispatchMemberGet(rs, "Source"); source.Type != VTString || source.Str != argued {
		t.Fatalf("the Open argument did not overwrite the Source property: %#v", source)
	}
}

// A fabricated (disconnected) recordset - fields appended, no connection, no
// Source - still opens empty rather than demanding a connection.
func TestRecordsetOpenWithoutSourceStillOpensAFabricatedRecordset(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	host.Server().SetRequestPath("/tests/test_adodb_source_property.asp")
	vm.SetHost(host)

	rs := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Recordset")})
	vm.dispatchNativeCall(rs.Num, "Fields.Append", []Value{NewString("col"), NewInteger(200)})
	vm.dispatchNativeCall(rs.Num, "Open", nil)

	if state := vm.dispatchMemberGet(rs, "State"); state.Type != VTInteger || state.Num != adStateOpen {
		t.Fatalf("fabricated recordset did not open: State = %#v", state)
	}
	if eof := vm.dispatchMemberGet(rs, "EOF"); eof.Type != VTBool || eof.Num == 0 {
		t.Fatalf("fabricated recordset should open at EOF: %#v", eof)
	}
}

// newSourcePropertyTestRecordset returns a VM with an open SQLite connection
// holding two known rows, and a recordset attached to that connection.
func newSourcePropertyTestRecordset(t *testing.T) (*VM, Value) {
	t.Helper()

	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_adodb_source_property.asp")
	vm.SetHost(host)

	conn := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	if conn.Type != VTNativeObject {
		t.Fatalf("expected connection object, got %#v", conn)
	}
	vm.dispatchMemberSet(conn.Num, "ConnectionString", NewString("sqlite:"+filepath.Join(rootDir, "source.db")))
	vm.dispatchNativeCall(conn.Num, "Open", nil)
	t.Cleanup(func() { vm.dispatchNativeCall(conn.Num, "Close", nil) })
	vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("CREATE TABLE people (id INTEGER PRIMARY KEY, name TEXT NOT NULL)")})
	vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("INSERT INTO people (id, name) VALUES (1, 'first'), (2, 'second')")})

	rs := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Recordset")})
	if rs.Type != VTNativeObject {
		t.Fatalf("expected recordset object, got %#v", rs)
	}
	vm.dispatchMemberSet(rs.Num, "ActiveConnection", conn)

	return vm, rs
}
