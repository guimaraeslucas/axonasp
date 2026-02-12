/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Test to directly query the Access database and inspect the sOrderBY field
 * for pages, specifically to debug list page detection.
 */
package server

import (
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"testing"
	"unicode/utf16"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func TestOLEFieldSOrderBY(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("OLE tests only run on Windows")
	}

	// Find the database file
	dbPath, err := filepath.Abs("../www/QuickerSite-test/db/data_jj2ar6as.mdb")
	if err != nil {
		t.Fatalf("Failed to resolve path: %v", err)
	}
	if _, err := os.Stat(dbPath); os.IsNotExist(err) {
		t.Skipf("Database not found at %s", dbPath)
	}

	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	if err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED); err != nil {
		t.Fatalf("CoInitializeEx failed: %v", err)
	}
	defer ole.CoUninitialize()

	// Create ADODB.Connection
	unknown, err := oleutil.CreateObject("ADODB.Connection")
	if err != nil {
		t.Fatalf("CreateObject ADODB.Connection failed: %v", err)
	}
	connDisp, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		t.Fatalf("QueryInterface failed: %v", err)
	}
	defer connDisp.Release()

	connStr := "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath
	_, err = oleutil.CallMethod(connDisp, "Open", connStr)
	if err != nil {
		// Try Jet provider if ACE is not available
		connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbPath
		_, err = oleutil.CallMethod(connDisp, "Open", connStr)
		if err != nil {
			t.Fatalf("Failed to open database: %v", err)
		}
	}
	defer oleutil.CallMethod(connDisp, "Close")

	t.Logf("Connected to database: %s", dbPath)

	// Query pages that have sOrderBY field populated - using a broader query first
	sqlText := "SELECT iId, sTitle, sOrderBY FROM tblPage WHERE iId=492"
	result, err := oleutil.CallMethod(connDisp, "Execute", sqlText)
	if err != nil {
		t.Fatalf("Execute failed: %v", err)
	}

	rsDisp := result.ToIDispatch()
	if rsDisp == nil {
		t.Fatal("Recordset is nil")
	}
	defer rsDisp.Release()

	// Check EOF
	eofResult, err := oleutil.GetProperty(rsDisp, "EOF")
	if err != nil {
		t.Fatalf("EOF check failed: %v", err)
	}
	if eofResult.Value() == true {
		t.Log("No record found for iId=492. Let's check all pages with sOrderBY...")
	} else {
		// Read sOrderBY value in multiple ways
		inspectFieldValue(t, rsDisp, "sOrderBY")
		inspectFieldValue(t, rsDisp, "sTitle")
		inspectFieldValue(t, rsDisp, "iId")
	}

	// Also query pages that have sOrderBY not empty/null - these are list pages
	t.Log("\n--- Pages with non-empty sOrderBY ---")
	sqlText2 := "SELECT iId, sTitle, sOrderBY FROM tblPage WHERE sOrderBY IS NOT NULL AND sOrderBY <> ''"
	result2, err := oleutil.CallMethod(connDisp, "Execute", sqlText2)
	if err != nil {
		t.Fatalf("Execute failed: %v", err)
	}
	rs2 := result2.ToIDispatch()
	if rs2 == nil {
		t.Fatal("Recordset 2 is nil")
	}
	defer rs2.Release()

	count := 0
	for {
		eofR, _ := oleutil.GetProperty(rs2, "EOF")
		if eofR.Value() == true {
			break
		}
		count++
		inspectFieldValue(t, rs2, "iId")
		inspectFieldValue(t, rs2, "sTitle")
		inspectFieldValue(t, rs2, "sOrderBY")
		t.Log("---")
		oleutil.CallMethod(rs2, "MoveNext")
		if count > 50 {
			t.Log("... (truncated)")
			break
		}
	}
	t.Logf("Total pages with non-empty sOrderBY: %d", count)
}

func inspectFieldValue(t *testing.T, rsDisp *ole.IDispatch, fieldName string) {
	t.Helper()

	// Get Fields collection
	fieldsResult, err := oleutil.GetProperty(rsDisp, "Fields")
	if err != nil {
		t.Logf("  [%s] Failed to get Fields: %v", fieldName, err)
		return
	}
	fieldsDisp := fieldsResult.ToIDispatch()
	if fieldsDisp == nil {
		t.Logf("  [%s] Fields is nil", fieldName)
		return
	}
	defer fieldsDisp.Release()

	// Get field by name
	fieldResult, err := oleutil.GetProperty(fieldsDisp, "Item", fieldName)
	if err != nil {
		t.Logf("  [%s] Field not found: %v", fieldName, err)
		return
	}
	fieldDisp := fieldResult.ToIDispatch()
	if fieldDisp == nil {
		t.Logf("  [%s] Field dispatch is nil", fieldName)
		return
	}
	defer fieldDisp.Release()

	// Get field type
	typeResult, err := oleutil.GetProperty(fieldDisp, "Type")
	if err != nil {
		t.Logf("  [%s] Failed to get Type: %v", fieldName, err)
	} else {
		t.Logf("  [%s] Field.Type = %v", fieldName, typeResult.Value())
	}

	// Get field size
	sizeResult, err := oleutil.GetProperty(fieldDisp, "DefinedSize")
	if err == nil {
		t.Logf("  [%s] Field.DefinedSize = %v", fieldName, sizeResult.Value())
	}

	// Get raw value as VARIANT
	valueResult, err := oleutil.GetProperty(fieldDisp, "Value")
	if err != nil {
		t.Logf("  [%s] Failed to get Value: %v", fieldName, err)
		return
	}

	// Print VARIANT details
	t.Logf("  [%s] VARIANT.VT = %d (0x%x)", fieldName, valueResult.VT, valueResult.VT)
	rawValue := valueResult.Value()
	t.Logf("  [%s] Value() type = %T", fieldName, rawValue)
	t.Logf("  [%s] Value() = %v", fieldName, rawValue)

	// If it's a string, print hex representation
	if s, ok := rawValue.(string); ok {
		t.Logf("  [%s] String len = %d", fieldName, len(s))
		if len(s) > 0 && len(s) < 200 {
			t.Logf("  [%s] String hex = %x", fieldName, []byte(s))
		}
	}

	// If it's []uint16, show the raw values
	if u16, ok := rawValue.([]uint16); ok {
		t.Logf("  [%s] []uint16 len = %d", fieldName, len(u16))
		if len(u16) > 0 && len(u16) < 100 {
			t.Logf("  [%s] []uint16 raw = %v", fieldName, u16)
			decoded := string(utf16.Decode(u16))
			t.Logf("  [%s] []uint16 decoded = %q", fieldName, decoded)
		}
	}

	// Also try direct access via rs(fieldName)
	directResult, err := oleutil.GetProperty(rsDisp, "Collect", fieldName)
	if err == nil {
		t.Logf("  [%s] Collect.VT = %d (0x%x)", fieldName, directResult.VT, directResult.VT)
		directValue := directResult.Value()
		t.Logf("  [%s] Collect type = %T, value = %v", fieldName, directValue, directValue)
		if s, ok := directValue.(string); ok {
			t.Logf("  [%s] Collect string len = %d, hex = %x", fieldName, len(s), []byte(s))
		}
	} else {
		t.Logf("  [%s] Collect failed: %v", fieldName, err)
	}

	// Normalize using our function
	normalized := normalizeOLEValue(rawValue)
	t.Logf("  [%s] normalizeOLEValue type = %T, value = %q", fieldName, normalized, fmt.Sprintf("%v", normalized))
}
