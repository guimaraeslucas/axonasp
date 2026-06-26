//go:build windows

package axonvm

import (
	"bytes"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"testing"
	"time"
)

func TestSTAWorkerInitializationAndShutdown(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	defer vm.Release()

	// Check that channels are initialized
	if vm.staTaskChan == nil {
		t.Fatal("expected staTaskChan to be initialized on Windows")
	}
	if vm.quitSTA == nil {
		t.Fatal("expected quitSTA to be initialized on Windows")
	}

	// Verify we can run a task on it
	run := false
	vm.runOnSTA(func() {
		run = true
	})
	if !run {
		t.Fatal("expected task to run on STA worker")
	}

	// Stop it and check that it stops
	vm.stopSTAWorker()

	if vm.staTaskChan != nil {
		t.Fatal("expected staTaskChan to be nil after stopping worker")
	}

	// Verify that runOnSTA executes directly now and doesn't block
	runDirect := false
	vm.runOnSTA(func() {
		runDirect = true
	})
	if !runDirect {
		t.Fatal("expected task to run directly after worker stops")
	}
}

func testOpenDatabase(dbPath string) (string, bool) {
	vm := NewVM(nil, nil, 5)
	defer vm.Release()

	connVal := vm.newADODBConnection()
	conn := vm.adodbConnectionItems[connVal.Num]
	if conn == nil {
		return "", false
	}
	conn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath

	var opened bool
	func() {
		defer func() {
			if r := recover(); r != nil {
				opened = false
			}
		}()
		vm.adodbConnectionOpen(conn)
		opened = (conn.state == adStateOpen)
	}()

	if !opened {
		return "", false
	}

	var selectOk bool
	// Try tblContact table
	func() {
		defer func() {
			if r := recover(); r != nil {
				selectOk = false
			}
		}()
		rsVal := vm.adodbConnectionExecute(conn, []Value{NewString("SELECT TOP 1 iId FROM tblContact")})
		rs := vm.adodbRecordsetItems[rsVal.Num]
		selectOk = (rs != nil)
	}()

	if selectOk {
		return "tblContact", true
	}

	// Try contact table
	func() {
		defer func() {
			if r := recover(); r != nil {
				selectOk = false
			}
		}()
		rsVal := vm.adodbConnectionExecute(conn, []Value{NewString("SELECT TOP 1 iId FROM contact")})
		rs := vm.adodbRecordsetItems[rsVal.Num]
		selectOk = (rs != nil)
	}()

	if selectOk {
		return "contact", true
	}

	// Try users table
	func() {
		defer func() {
			if r := recover(); r != nil {
				selectOk = false
			}
		}()
		rsVal := vm.adodbConnectionExecute(conn, []Value{NewString("SELECT TOP 1 id FROM users")})
		rs := vm.adodbRecordsetItems[rsVal.Num]
		selectOk = (rs != nil)
	}()

	if selectOk {
		return "users", true
	}

	return "", false
}

func findAccessDatabase(t *testing.T) (string, string) {
	repoRoot, err := filepath.Abs("..")
	if err != nil {
		t.Fatalf("resolve repo root: %v", err)
	}

	// Check standard location
	originalPath := filepath.Join(repoRoot, "www", "asplite-test", "db", "sample.mdb")
	if _, err := os.Stat(originalPath); err == nil {
		if tbl, ok := testOpenDatabase(originalPath); ok {
			return originalPath, tbl
		}
	}

	// Search in www/temp/uploads/
	uploadsDir := filepath.Join(repoRoot, "www", "temp", "uploads")
	files, err := os.ReadDir(uploadsDir)
	if err == nil {
		for _, f := range files {
			if strings.HasSuffix(strings.ToLower(f.Name()), ".mdb") {
				path := filepath.Join(uploadsDir, f.Name())
				if tbl, ok := testOpenDatabase(path); ok {
					return path, tbl
				}
			}
		}
	}
	return "", ""
}

func TestADODBLoopConnectionNoMemoryLeak(t *testing.T) {
	dbPath, tblName := findAccessDatabase(t)
	if dbPath == "" {
		t.Skip("No compatible Access database found for testing memory leaks")
	}

	var selectQuery string
	if tblName == "tblContact" {
		selectQuery = "SELECT TOP 1 iId FROM tblContact"
	} else if tblName == "contact" {
		selectQuery = "SELECT TOP 1 iId FROM contact"
	} else {
		selectQuery = "SELECT TOP 1 id FROM users"
	}

	warmUp := 20
	runIterations := func(iterations int) {
		for range iterations {
			func() {
				vm := NewVM(nil, nil, 5)
				defer vm.Release()

				connVal := vm.newADODBConnection()
				conn := vm.adodbConnectionItems[connVal.Num]
				if conn == nil {
					t.Fatal("expected ADODB connection instance")
				}
				conn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath
				vm.adodbConnectionOpen(conn)
				if conn.state != adStateOpen {
					t.Fatalf("expected open connection, state=%d", conn.state)
				}

				// Execute a simple query
				rsVal := vm.adodbConnectionExecute(conn, []Value{NewString(selectQuery)})
				rs := vm.adodbRecordsetItems[rsVal.Num]
				if rs == nil {
					t.Fatal("expected Recordset")
				}
			}()
		}
	}

	// Warm-up
	runIterations(warmUp)
	runtime.GC()
	runtime.GC()
	time.Sleep(200 * time.Millisecond)

	var ms1 runtime.MemStats
	runtime.ReadMemStats(&ms1)

	// Run loop
	runIterations(100)

	runtime.GC()
	runtime.GC()
	time.Sleep(200 * time.Millisecond)

	var ms2 runtime.MemStats
	runtime.ReadMemStats(&ms2)

	heapDiff := int64(ms2.HeapAlloc) - int64(ms1.HeapAlloc)
	t.Logf("HeapAlloc after warm-up: %d, after loop: %d, diff: %d", ms1.HeapAlloc, ms2.HeapAlloc, heapDiff)

	// Heap growth should be minimal (less than 8MB)
	if heapDiff > 8*1024*1024 {
		t.Fatalf("Possible memory leak: HeapAlloc grew by %d bytes", heapDiff)
	}
}

func TestASPStressConcurrentReadWrite(t *testing.T) {
	dbPath, tblName := findAccessDatabase(t)
	if dbPath == "" {
		t.Skip("No compatible Access database found for concurrent stress testing")
	}

	tempDir := t.TempDir()
	tempDbPath := filepath.Join(tempDir, "stress_sample.mdb")

	// Copy the database file to tempDbPath
	data, err := os.ReadFile(dbPath)
	if err != nil {
		t.Fatalf("failed to read source database: %v", err)
	}
	err = os.WriteFile(tempDbPath, data, 0666)
	if err != nil {
		t.Fatalf("failed to write temp mdb: %v", err)
	}

	// Diagnostic: print column names
	{
		vm := NewVM(nil, nil, 5)
		defer vm.Release()
		connVal := vm.newADODBConnection()
		conn := vm.adodbConnectionItems[connVal.Num]
		conn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tempDbPath
		vm.adodbConnectionOpen(conn)
		if conn.state == adStateOpen {
			rsVal := vm.adodbConnectionExecute(conn, []Value{NewString("SELECT TOP 1 * FROM " + tblName)})
			rs := vm.adodbRecordsetItems[rsVal.Num]
			if rs != nil {
				t.Logf("DIAGNOSTIC: Table %s columns: %v", tblName, rs.columns)
			}
		}
	}

	var source string
	if tblName == "tblContact" {
		source = `<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_PATH_PLACEHOLDER"
customerID = 73
On Error Resume Next
Set custRS = conn.Execute("SELECT TOP 1 iId FROM tblCustomer")
If Err.Number = 0 And Not custRS.EOF Then
    val = custRS("iId").Value
    If Not IsNull(val) And val <> "" Then
        customerID = val
    End If
End If
custRS.Close
Err.Clear
On Error GoTo 0
Set rs = conn.Execute("SELECT iId, sEmail FROM tblContact")
count = 0
Do While Not rs.EOF
    count = count + 1
    rs.MoveNext
Loop
rs.Close
conn.Execute "INSERT INTO tblContact (sEmail, sNickName, iCustomerID) VALUES ('stress@test.com', 'stress nickname', " & customerID & ")"
conn.Execute "UPDATE tblContact SET sNickName = 'stress nickname updated' WHERE sEmail = 'stress@test.com'"
conn.Close
%>`
	} else if tblName == "contact" {
		source = `<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_PATH_PLACEHOLDER"
Set rs = conn.Execute("SELECT iId, sText FROM contact WHERE iCountryID = 1")
count = 0
Do While Not rs.EOF
    count = count + 1
    rs.MoveNext
Loop
rs.Close
conn.Execute "INSERT INTO contact (sText, iNumber, iCountryID) VALUES ('stress test', " & count & ", 1)"
conn.Execute "UPDATE contact SET sText = 'stress test updated' WHERE iNumber = " & count
conn.Close
%>`
	} else {
		source = `<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_PATH_PLACEHOLDER"
Set rs = conn.Execute("SELECT id, name FROM users")
count = 0
Do While Not rs.EOF
    count = count + 1
    rs.MoveNext
Loop
rs.Close
conn.Execute "INSERT INTO users (name, email, age) VALUES ('stress test', 'stress@test.com', 30)"
conn.Execute "UPDATE users SET name = 'stress test updated' WHERE email = 'stress@test.com'"
conn.Close
%>`
	}
	source = strings.ReplaceAll(source, "DB_PATH_PLACEHOLDER", strings.ReplaceAll(tempDbPath, "\\", "\\\\"))

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("failed to compile stress ASP script: %v", err)
	}
	bytecode := compiler.Bytecode()
	constants := compiler.Constants()
	globalsCount := compiler.GlobalsCount()

	var wg sync.WaitGroup
	errs := make(chan error, 30)

	// Simulating vm_pool_size >= 10 with 15 concurrent execution paths
	for i := range 15 {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()
			vm := NewVM(bytecode, constants, globalsCount)
			defer vm.Release()

			host := NewMockHost()
			var output bytes.Buffer
			host.SetOutput(&output)
			vm.SetHost(host)

			err := vm.Run()
			if err != nil {
				errStr := err.Error()
				// Ignore database locked or sharing violation errors which are expected
				// from standard Microsoft Access Jet/ACE driver when writing concurrently.
				if strings.Contains(errStr, "locked") || strings.Contains(errStr, "sharing violation") || strings.Contains(errStr, "state") || strings.Contains(errStr, "bloqueado") || strings.Contains(errStr, "compartilhamento") {
					return
				}
				errs <- err
			}
		}(i)
	}
	wg.Wait()
	close(errs)

	for err := range errs {
		t.Errorf("concurrent execution crash or unexpected error: %v", err)
	}
}
