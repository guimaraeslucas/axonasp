package server

import (
	"fmt"
	"net/http/httptest"
	"os"
	"path/filepath"
	"runtime"
	"testing"
	"time"
)

func TestAccessOLERecordsetUpdate(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Access OLEDB tests require Windows")
	}

	rootDir, err := testWWWRoot()
	if err != nil {
		t.Fatalf("failed to resolve www root: %v", err)
	}

	ctx := NewExecutionContext(httptest.NewRecorder(), httptest.NewRequest("GET", "http://localhost", nil), "TESTSESSION", 10*time.Second)
	ctx.RootDir = rootDir
	ctx.Server.SetRootDir(rootDir)

	conn := NewADODBConnection(ctx)
	conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ctx.Server_MapPath("db/sample.mdb")
	conn.CallMethod("open")
	if toInt(conn.GetProperty("state")) != 1 {
		t.Skip("Access OLEDB provider not available or connection failed")
	}
	defer conn.CallMethod("close")

	rs := NewADODBRecordset(ctx)
	rs.SetProperty("cursortype", 1)
	rs.SetProperty("locktype", 3)
	rs.SetProperty("activeconnection", conn)

	rs.CallMethod("open", "select * from contact where 1=2")
	rs.CallMethod("addnew")

	insertText := fmt.Sprintf("TEST_SAVE_%d", time.Now().UnixNano())
	rs.SetProperty("sText", insertText)
	rs.SetProperty("iNumber", 93)
	rs.SetProperty("dDate", time.Date(1973, time.July, 15, 0, 0, 0, 0, time.Local))
	rs.SetProperty("iCountryID", 143)
	rs.CallMethod("update")

	newID := toInt(rs.CallMethod("", "iId"))
	if newID == 0 {
		t.Fatalf("expected non-zero iId after insert")
	}

	rs.CallMethod("close")

	rs = NewADODBRecordset(ctx)
	rs.SetProperty("cursortype", 1)
	rs.SetProperty("locktype", 3)
	rs.SetProperty("activeconnection", conn)
	rs.CallMethod("open", fmt.Sprintf("select * from contact where iId=%d", newID))

	updateText := fmt.Sprintf("TEST_UPDATE_%d", time.Now().UnixNano())
	rs.SetProperty("sText", updateText)
	rs.SetProperty("iNumber", 94)
	rs.SetProperty("dDate", time.Date(1973, time.July, 16, 0, 0, 0, 0, time.Local))
	rs.SetProperty("iCountryID", 143)
	rs.CallMethod("update")

	rs.CallMethod("close")

	rs = NewADODBRecordset(ctx)
	rs.SetProperty("activeconnection", conn)
	rs.CallMethod("open", fmt.Sprintf("select * from contact where iId=%d", newID))

	readText := fmt.Sprintf("%v", rs.CallMethod("", "sText"))
	if readText != updateText {
		t.Fatalf("expected updated sText, got %q", readText)
	}
}

func TestAccessOLEFieldsEnumeration(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Access OLEDB tests require Windows")
	}

	rootDir, err := testWWWRoot()
	if err != nil {
		t.Fatalf("failed to resolve www root: %v", err)
	}

	ctx := NewExecutionContext(httptest.NewRecorder(), httptest.NewRequest("GET", "http://localhost", nil), "TESTSESSION", 10*time.Second)
	ctx.RootDir = rootDir
	ctx.Server.SetRootDir(rootDir)

	conn := NewADODBConnection(ctx)
	conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ctx.Server_MapPath("db/sample.mdb")
	conn.CallMethod("open")
	if toInt(conn.GetProperty("state")) != 1 {
		t.Skip("Access OLEDB provider not available or connection failed")
	}
	defer conn.CallMethod("close")

	rs := NewADODBRecordset(ctx)
	rs.SetProperty("cursortype", 1)
	rs.SetProperty("locktype", 3)
	rs.SetProperty("activeconnection", conn)
	rs.CallMethod("open", "select * from contact where 1=2")
	defer rs.CallMethod("close")

	fields := rs.GetProperty("fields")
	enum, ok := fields.(interface{ Enumeration() []interface{} })
	if !ok {
		t.Fatalf("expected fields enumeration, got %T", fields)
	}

	items := enum.Enumeration()
	if len(items) == 0 {
		t.Fatalf("expected fields enumeration to return items")
	}
}

func testWWWRoot() (string, error) {
	wd, err := os.Getwd()
	if err != nil {
		return "", err
	}

	root := filepath.Clean(filepath.Join(wd, "..", "www"))
	if _, statErr := os.Stat(root); statErr == nil {
		return root, nil
	}

	root = filepath.Clean(filepath.Join(wd, "www"))
	if _, statErr := os.Stat(root); statErr == nil {
		return root, nil
	}

	return "", fmt.Errorf("www root not found from %s", wd)
}
