//go:build !wasm

package axonvm

import (
	"bytes"
	"context"
	"database/sql"
	"database/sql/driver"
	"errors"
	"io"
	"path/filepath"
	"strings"
	"sync"
	"testing"
)

type mockIdentityDriverState struct {
	mu               sync.Mutex
	identityColName  string
	scopeIdentityVal any
	atAtIdentityVal  any
	executedQueries  []string
	txExecCount      int
	txQueryCount     int
	connExecCount    int
	connQueryCount   int
}

var (
	mockIdentityDriversMu sync.Mutex
	mockIdentityStates    = make(map[string]*mockIdentityDriverState)
)

func init() {
	sql.Register("adodb-mock-identity", &mockIdentityDriver{})
}

type mockIdentityDriver struct{}

func (d *mockIdentityDriver) Open(name string) (driver.Conn, error) {
	mockIdentityDriversMu.Lock()
	state := mockIdentityStates[name]
	mockIdentityDriversMu.Unlock()
	if state == nil {
		return nil, errors.New("unknown mock identity dsn")
	}
	return &mockIdentityConn{state: state}, nil
}

type mockIdentityConn struct {
	state *mockIdentityDriverState
	inTx  bool
}

func (c *mockIdentityConn) Prepare(query string) (driver.Stmt, error) {
	return &mockIdentityStmt{conn: c, query: query}, nil
}

func (c *mockIdentityConn) Close() error { return nil }

func (c *mockIdentityConn) Begin() (driver.Tx, error) {
	c.inTx = true
	return &mockIdentityTx{conn: c}, nil
}

func (c *mockIdentityConn) ExecContext(_ context.Context, query string, _ []driver.NamedValue) (driver.Result, error) {
	c.state.mu.Lock()
	c.state.executedQueries = append(c.state.executedQueries, query)
	if c.inTx {
		c.state.txExecCount++
	} else {
		c.state.connExecCount++
	}
	c.state.mu.Unlock()
	return &mockIdentityResult{}, nil
}

func (c *mockIdentityConn) QueryContext(_ context.Context, query string, _ []driver.NamedValue) (driver.Rows, error) {
	c.state.mu.Lock()
	c.state.executedQueries = append(c.state.executedQueries, query)
	if c.inTx {
		c.state.txQueryCount++
	} else {
		c.state.connQueryCount++
	}
	idCol := c.state.identityColName
	scopeVal := c.state.scopeIdentityVal
	atAtVal := c.state.atAtIdentityVal
	c.state.mu.Unlock()

	upper := strings.ToUpper(query)
	switch {
	case strings.Contains(upper, "SYS.IDENTITY_COLUMNS"):
		return &mockIdentityRows{
			cols: []string{"name"},
			data: [][]driver.Value{{idCol}},
		}, nil
	case strings.Contains(upper, "SCOPE_IDENTITY()"):
		return &mockIdentityRows{
			cols: []string{"NewID"},
			data: [][]driver.Value{{scopeVal}},
		}, nil
	case strings.Contains(upper, "@@IDENTITY"):
		return &mockIdentityRows{
			cols: []string{"NewID"},
			data: [][]driver.Value{{atAtVal}},
		}, nil
	case strings.Contains(upper, "LAST_INSERT_ID()"), strings.Contains(upper, "LAST_INSERT_ROWID()"):
		return &mockIdentityRows{
			cols: []string{"NewID"},
			data: [][]driver.Value{{scopeVal}},
		}, nil
	default:
		return &mockIdentityRows{
			cols: []string{idCol, "Name"},
			data: nil,
		}, nil
	}
}

type mockIdentityTx struct {
	conn *mockIdentityConn
}

func (t *mockIdentityTx) Commit() error {
	t.conn.inTx = false
	return nil
}

func (t *mockIdentityTx) Rollback() error {
	t.conn.inTx = false
	return nil
}

type mockIdentityStmt struct {
	conn  *mockIdentityConn
	query string
}

func (s *mockIdentityStmt) Close() error  { return nil }
func (s *mockIdentityStmt) NumInput() int { return 0 }
func (s *mockIdentityStmt) Exec(_ []driver.Value) (driver.Result, error) {
	return s.conn.ExecContext(context.Background(), s.query, nil)
}
func (s *mockIdentityStmt) Query(_ []driver.Value) (driver.Rows, error) {
	return s.conn.QueryContext(context.Background(), s.query, nil)
}

type mockIdentityResult struct{}

func (r *mockIdentityResult) LastInsertId() (int64, error) {
	return -1, errors.New("LastInsertId is not supported")
}

func (r *mockIdentityResult) RowsAffected() (int64, error) {
	return 1, nil
}

type mockIdentityRows struct {
	cols []string
	data [][]driver.Value
	idx  int
}

func (r *mockIdentityRows) Columns() []string { return r.cols }
func (r *mockIdentityRows) Close() error      { return nil }
func (r *mockIdentityRows) Next(dest []driver.Value) error {
	if r.idx >= len(r.data) {
		return io.EOF
	}
	copy(dest, r.data[r.idx])
	r.idx++
	return nil
}

// TestADODBIdentityRetrievalSQLiteEndToEnd verifies Recordset.Update after AddNew populates
// auto-increment identities for standard, custom-named, and pre-fetched field proxy reads.
func TestADODBIdentityRetrievalSQLiteEndToEnd(t *testing.T) {
	dbPath := filepath.Join(t.TempDir(), "identity_e2e.db")
	source := `<%
Dim conn, rs, fId
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "sqlite:" & Server.MapPath("/identity_e2e.db")

conn.Execute "CREATE TABLE items (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL)"
conn.Execute "CREATE TABLE custom_items (item_pk INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL)"
conn.Execute "CREATE TABLE manual_items (id INTEGER, title TEXT NOT NULL)"
conn.Execute "INSERT INTO manual_items (id, title) VALUES (1, 'existing')"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT * FROM items WHERE 1=2", conn
Set fId = rs.Fields("id")
rs.AddNew
rs.Fields("name").Value = "first"
rs.Update
Response.Write "std=" & rs.Fields("id").Value & "|" & rs("id") & "|" & fId.Value & "|"
rs.AddNew
rs("name") = "second"
rs.Update
Response.Write rs.Fields("id").Value & "|"
rs.Close

rs.Open "SELECT * FROM custom_items WHERE 1=2", conn
rs.AddNew
rs.Fields("title").Value = "custom"
rs.Update
Response.Write "custom=" & rs.Fields("item_pk").Value & "|" & rs("item_pk") & "|"
rs.Close

rs.Open "SELECT * FROM manual_items WHERE 1=2", conn
rs.AddNew
rs("id") = 99
rs("title") = "manual"
rs.Update
Response.Write "manual=" & rs.Fields("id").Value
rs.Close

conn.Close
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	host.Server().SetRootDir(filepath.Dir(dbPath))
	host.Server().SetRequestPath("/test_identity.asp")
	var out bytes.Buffer
	host.SetOutput(&out)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("run failed: %v", err)
	}
	host.Response().Flush()

	want := "std=1|1|1|2|custom=1|1|manual=99"
	if out.String() != want {
		t.Fatalf("got %q, want %q", out.String(), want)
	}
}

// TestADODBIdentityRetrievalMSSQLMockAndTx verifies SQL Server fallback to SCOPE_IDENTITY()
// when LastInsertId() returns an error, both outside and inside a transaction.
func TestADODBIdentityRetrievalMSSQLMockAndTx(t *testing.T) {
	dsn := "mssql-mock-identity-test"
	state := &mockIdentityDriverState{
		identityColName:  "UserID",
		scopeIdentityVal: []byte("42"),
		atAtIdentityVal:  []byte("99"),
	}
	mockIdentityDriversMu.Lock()
	mockIdentityStates[dsn] = state
	mockIdentityDriversMu.Unlock()
	defer func() {
		mockIdentityDriversMu.Lock()
		delete(mockIdentityStates, dsn)
		mockIdentityDriversMu.Unlock()
	}()

	db, err := sql.Open("adodb-mock-identity", dsn)
	if err != nil {
		t.Fatalf("sql.Open: %v", err)
	}
	defer db.Close()

	dbConn, err := db.Conn(context.Background())
	if err != nil {
		t.Fatalf("db.Conn: %v", err)
	}
	defer dbConn.Close()

	vm := NewVM(nil, nil, 5)
	connVal := vm.newADODBConnection()
	conn := vm.adodbConnectionItems[connVal.Num]
	conn.db = db
	conn.dbConn = dbConn
	conn.dbDriver = "mssql"
	conn.provider = "SQLOLEDB"
	conn.state = adStateOpen

	rsVal := vm.newADODBRecordset()
	rs := vm.adodbRecordsetItems[rsVal.Num]
	rs.activeConnection = connVal.Num
	rs.source = "SELECT * FROM Users WHERE 1=2"
	vm.adodbRecordsetOpen(rs, rs.source, conn, []Value{NewString(rs.source)})

	userIDField := vm.dispatchADODBRecordsetMethod(rs, "Fields", []Value{NewString("UserID")})
	nameField := vm.dispatchADODBRecordsetMethod(rs, "Fields", []Value{NewString("Name")})

	vm.adodbRecordsetAddNew(rs, nil)
	if !vm.dispatchADODBFieldPropertySet(nameField.Num, "Value", NewString("Alice")) {
		t.Fatal("expected field Value set to succeed")
	}
	vm.adodbRecordsetUpdate(rs, nil)
	if vm.lastError != nil {
		t.Fatalf("unexpected update error: %v", vm.lastError)
	}

	gotVal, ok := vm.dispatchADODBFieldPropertyGet(userIDField.Num, "Value")
	if !ok || gotVal.Type != VTInteger || gotVal.Num != 42 {
		t.Fatalf("expected UserID=42 after Update, got %#v (ok=%v)", gotVal, ok)
	}

	// Verify transaction context runs both INSERT and SCOPE_IDENTITY() on conn.tx.
	vm.adodbConnectionBeginTrans(conn)
	if conn.tx == nil {
		t.Fatal("expected active transaction")
	}
	state.mu.Lock()
	state.scopeIdentityVal = nil // force @@IDENTITY fallback inside tx
	state.atAtIdentityVal = int64(77)
	state.mu.Unlock()

	vm.adodbRecordsetAddNew(rs, nil)
	vm.dispatchADODBFieldPropertySet(nameField.Num, "Value", NewString("Bob"))
	vm.adodbRecordsetUpdate(rs, nil)
	if vm.lastError != nil {
		t.Fatalf("unexpected tx update error: %v", vm.lastError)
	}
	vm.adodbConnectionCommitTrans(conn)

	gotTxVal, ok := vm.dispatchADODBFieldPropertyGet(userIDField.Num, "Value")
	if !ok || gotTxVal.Type != VTInteger || gotTxVal.Num != 77 {
		t.Fatalf("expected UserID=77 after tx Update, got %#v (ok=%v)", gotTxVal, ok)
	}

	state.mu.Lock()
	defer state.mu.Unlock()
	if state.txExecCount != 1 || state.txQueryCount < 2 {
		t.Fatalf("expected tx exec/query counts (1, >=2), got (%d, %d)", state.txExecCount, state.txQueryCount)
	}
}

// TestADODBIdentityFallbackQueriesByDriver verifies provider-specific identity SQL queries
// when LastInsertId returns an error on MySQL and SQLite connections.
func TestADODBIdentityFallbackQueriesByDriver(t *testing.T) {
	for _, tc := range []struct {
		driver       string
		wantSubstr   string
		expectedID   int64
		returnedBody any
	}{
		{driver: "mysql", wantSubstr: "LAST_INSERT_ID()", expectedID: 15, returnedBody: int64(15)},
		{driver: "sqlite", wantSubstr: "last_insert_rowid()", expectedID: 23, returnedBody: int64(23)},
		{driver: "mssql", wantSubstr: "SCOPE_IDENTITY()", expectedID: 31, returnedBody: []byte("31")},
	} {
		t.Run(tc.driver, func(t *testing.T) {
			dsn := "mock-driver-fallback-" + tc.driver
			state := &mockIdentityDriverState{
				identityColName:  "id",
				scopeIdentityVal: tc.returnedBody,
			}
			mockIdentityDriversMu.Lock()
			mockIdentityStates[dsn] = state
			mockIdentityDriversMu.Unlock()
			defer func() {
				mockIdentityDriversMu.Lock()
				delete(mockIdentityStates, dsn)
				mockIdentityDriversMu.Unlock()
			}()

			db, err := sql.Open("adodb-mock-identity", dsn)
			if err != nil {
				t.Fatalf("sql.Open: %v", err)
			}
			defer db.Close()

			dbConn, err := db.Conn(context.Background())
			if err != nil {
				t.Fatalf("db.Conn: %v", err)
			}
			defer dbConn.Close()

			vm := NewVM(nil, nil, 5)
			connVal := vm.newADODBConnection()
			conn := vm.adodbConnectionItems[connVal.Num]
			conn.db = db
			conn.dbConn = dbConn
			conn.dbDriver = tc.driver
			conn.state = adStateOpen

			id, ok := vm.adodbExecWriteback(conn, "INSERT INTO items (name) VALUES ('x')", "Test", true)
			if !ok || id != tc.expectedID {
				t.Fatalf("expected id=%d ok=true, got id=%d ok=%v", tc.expectedID, id, ok)
			}
			state.mu.Lock()
			queries := strings.Join(state.executedQueries, "\n")
			state.mu.Unlock()
			if !strings.Contains(queries, tc.wantSubstr) {
				t.Fatalf("expected query containing %q, got:\n%s", tc.wantSubstr, queries)
			}
		})
	}
}
