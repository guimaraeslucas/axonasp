package axonvm

import "testing"

// ADO lets a connection string be assigned to ActiveConnection, on a Recordset
// or on a Command, and opens an implicit connection from it. Accepting only a
// native object dropped the assignment silently, so the recordset ended up with
// no connection and the later Open() failed with "no active connection" -
// blaming the Open rather than the assignment that did nothing.
//
// Classic ASP code does this routinely:
//
//	set rf = Server.CreateObject("ADODB.Recordset")
//	rf.ActiveConnection = postgresConn   ' a connection string
//	rf.source = "select ..."
//	rf.open()
func TestRecordsetActiveConnectionAcceptsAConnectionString(t *testing.T) {
	vm := newADODBTestVM()
	rs := &adodbRecordset{}

	const connStr = "Driver=postgres;Server=localhost;Database=demo"
	if !vm.dispatchADODBRecordsetPropertySet(rs, "ActiveConnection", NewString(connStr)) {
		t.Fatal("the setter did not claim the ActiveConnection property")
	}
	if rs.activeConnection == 0 {
		t.Fatal("assigning a connection string left the recordset with no " +
			"connection; Open() would fail with \"no active connection\"")
	}
	conn, ok := vm.adodbConnectionItems[rs.activeConnection]
	if !ok {
		t.Fatalf("activeConnection %d names no registered connection", rs.activeConnection)
	}
	if conn.connectionString != connStr {
		t.Errorf("connection string = %q, want %q", conn.connectionString, connStr)
	}
}

func TestCommandActiveConnectionAcceptsAConnectionString(t *testing.T) {
	vm := newADODBTestVM()
	cmd := &adodbCommand{}

	const connStr = "Driver=postgres;Server=localhost;Database=demo"
	vm.dispatchADODBCommandPropertySet(cmd, "ActiveConnection", NewString(connStr))
	if cmd.activeConnection == 0 {
		t.Fatal("assigning a connection string left the command with no connection")
	}
	if conn := vm.adodbConnectionItems[cmd.activeConnection]; conn == nil ||
		conn.connectionString != connStr {
		t.Errorf("command connection string not recorded")
	}
}

// Nothing and an empty string must keep doing what they did before: an object
// assignment is how a recordset is attached, and assigning Nothing to
// ActiveConnection after opening is a common idiom that previously changed
// nothing. Making that detach is a separate question and not part of this fix.
func TestRecordsetActiveConnectionIgnoresEmptyValues(t *testing.T) {
	vm := newADODBTestVM()
	for name, val := range map[string]Value{
		"empty string": NewString(""),
		"blank string": NewString("   "),
		"Nothing":      {Type: VTNothing},
		"Empty":        {Type: VTEmpty},
	} {
		rs := &adodbRecordset{activeConnection: 41}
		vm.dispatchADODBRecordsetPropertySet(rs, "ActiveConnection", val)
		if rs.activeConnection != 41 {
			t.Errorf("%s changed activeConnection to %d; it should leave it alone",
				name, rs.activeConnection)
		}
	}
}

func newADODBTestVM() *VM {
	return &VM{
		adodbConnectionItems: make(map[int64]*adodbConnection),
		nextDynamicNativeID:  1,
	}
}
