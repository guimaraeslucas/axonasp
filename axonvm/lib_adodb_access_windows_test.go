//go:build windows

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
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"time"
)

func TestADODBAccessJoinRecordsetOpen(t *testing.T) {
	repoRoot, err := filepath.Abs("..")
	if err != nil {
		t.Fatalf("resolve repo root: %v", err)
	}
	dbPath := filepath.Join(repoRoot, "www", "asplite-test", "db", "sample.mdb")
	if _, err := os.Stat(dbPath); err != nil {
		t.Skipf("sample database not available: %v", err)
	}

	vm := NewVM(nil, nil, 5)
	defer vm.CleanupRequestResources()

	connVal := vm.newADODBConnection()
	conn := vm.adodbConnectionItems[connVal.Num]
	if conn == nil {
		t.Fatal("expected ADODB connection instance")
	}
	conn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath
	vm.adodbConnectionOpen(conn)
	if conn.state != adStateOpen {
		t.Fatalf("expected Access connection open, state=%d", conn.state)
	}

	rsVal := vm.newADODBRecordset()
	rs := vm.adodbRecordsetItems[rsVal.Num]
	if rs == nil {
		t.Fatal("expected ADODB recordset instance")
	}
	rs.activeConnection = connVal.Num
	rs.cursorType = adOpenKeyset
	rs.lockType = adLockOptimistic

	sqlText := "SELECT contact.iId, contact.sText, contact.iNumber, contact.dDate, country.sText AS countryName FROM contact LEFT JOIN country ON contact.iCountryID = country.iId ORDER BY contact.iId"
	vm.adodbRecordsetOpen(rs, sqlText, conn, []Value{NewString(sqlText)})

	if rs.state != adStateOpen {
		t.Fatalf("expected recordset open, state=%d", rs.state)
	}
	recordCount := vm.dispatchMemberGet(rsVal, "RecordCount")
	if recordCount.Type != VTInteger || recordCount.Num <= 0 {
		t.Fatalf("unexpected record count: %#v", recordCount)
	}
	vm.dispatchMemberSet(rsVal.Num, "AbsolutePosition", NewInteger(10))
	fieldVal := vm.dispatchMemberGet(rsVal, "countryName")
	if fieldVal.Type == VTEmpty {
		t.Fatal("expected non-empty field access after positioning")
	}

	rowsRead := 0
	fieldsVal := vm.dispatchMemberGet(rsVal, "Fields")
	fieldCountVal := vm.dispatchMemberGet(fieldsVal, "Count")
	if fieldCountVal.Type != VTInteger || fieldCountVal.Num <= 0 {
		t.Fatalf("unexpected fields count: %#v", fieldCountVal)
	}
	for rowsRead < 20 {
		eof := vm.dispatchMemberGet(rsVal, "EOF")
		if eof.Type == VTBool && eof.Num != 0 {
			break
		}
		for i := int64(0); i < fieldCountVal.Num; i++ {
			fieldVal, ok := vm.dispatchADODBFieldsCollectionMethod(fieldsVal.Num, "", []Value{NewInteger(i)})
			if !ok || fieldVal.Type != VTNativeObject {
				t.Fatalf("expected field proxy for index %d, got %#v", i, fieldVal)
			}
			nameVal := vm.dispatchMemberGet(fieldVal, "Name")
			valueVal := vm.dispatchMemberGet(fieldVal, "Value")
			if nameVal.Type != VTString || nameVal.Str == "" {
				t.Fatalf("unexpected field name at index %d: %#v", i, nameVal)
			}
			if valueVal.Type == VTEmpty && nameVal.Str == "iId" {
				t.Fatal("expected non-empty iId field during iteration")
			}
		}
		vm.adodbRecordsetMoveNext(rs)
		rowsRead++
	}
	if rowsRead == 0 {
		t.Fatal("expected to iterate at least one row")
	}
}

func TestADODBAccessRecordsetUpdatePersists(t *testing.T) {
	repoRoot, err := filepath.Abs("..")
	if err != nil {
		t.Fatalf("resolve repo root: %v", err)
	}
	dbPath := filepath.Join(repoRoot, "www", "asplite-test", "db", "sample.mdb")
	if _, err := os.Stat(dbPath); err != nil {
		t.Skipf("sample database not available: %v", err)
	}

	vm := NewVM(nil, nil, 5)
	defer vm.CleanupRequestResources()

	connVal := vm.newADODBConnection()
	conn := vm.adodbConnectionItems[connVal.Num]
	if conn == nil {
		t.Fatal("expected ADODB connection instance")
	}
	conn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath
	vm.adodbConnectionOpen(conn)
	if conn.state != adStateOpen {
		t.Fatalf("expected Access connection open, state=%d", conn.state)
	}

	firstRowVal := vm.adodbConnectionExecute(conn, []Value{NewString("select top 1 * from contact order by iId")})
	firstRow := vm.adodbRecordsetItems[firstRowVal.Num]
	if firstRow == nil || firstRow.recordCount == 0 {
		t.Fatal("expected at least one contact row in sample database")
	}
	targetID := vm.asInt(vm.dispatchMemberGet(firstRowVal, "iId"))
	if targetID == 0 {
		t.Fatalf("expected non-zero contact iId, got %d", targetID)
	}

	loadRow := func() (Value, Value) {
		rsVal := vm.newADODBRecordset()
		rs := vm.adodbRecordsetItems[rsVal.Num]
		if rs == nil {
			t.Fatal("expected ADODB recordset instance")
		}
		rs.activeConnection = connVal.Num
		rs.cursorType = adOpenKeyset
		rs.lockType = adLockOptimistic
		sqlText := "select * from contact where iId=" + strconv.Itoa(targetID)
		vm.adodbRecordsetOpen(rs, sqlText, conn, []Value{NewString(sqlText)})
		if rs.state != adStateOpen || rs.recordCount != 1 {
			t.Fatalf("expected one contact row, state=%d count=%d", rs.state, rs.recordCount)
		}
		return vm.dispatchMemberGet(rsVal, "sText"), vm.dispatchMemberGet(rsVal, "iNumber")
	}

	originalText, originalNumber := loadRow()
	defer func() {
		restoreSQL := "UPDATE [contact] SET [sText] = '" + strings.ReplaceAll(originalText.String(), "'", "''") + "', [iNumber] = " + strconv.FormatInt(int64(vm.asInt(originalNumber)), 10) + " WHERE iId=" + strconv.Itoa(targetID)
		_, _ = vm.adodbExecWriteback(conn, restoreSQL, "TestADODBAccessRecordsetUpdatePersists", false)
	}()

	rsVal := vm.newADODBRecordset()
	rs := vm.adodbRecordsetItems[rsVal.Num]
	if rs == nil {
		t.Fatal("expected ADODB recordset instance")
	}
	rs.activeConnection = connVal.Num
	rs.cursorType = adOpenKeyset
	rs.lockType = adLockOptimistic
	sqlText := "select * from contact where iId=" + strconv.Itoa(targetID)
	vm.adodbRecordsetOpen(rs, sqlText, conn, []Value{NewString(sqlText)})
	if rs.state != adStateOpen || rs.recordCount != 1 {
		t.Fatalf("expected one contact row, state=%d count=%d", rs.state, rs.recordCount)
	}

	updatedText := NewString("AxonASP update persistence")
	updatedNumber := NewInteger(424242)
	textField := vm.newADODBFieldProxy(rs, "sText")
	numberField := vm.newADODBFieldProxy(rs, "iNumber")
	vm.dispatchMemberSet(textField.Num, "Value", updatedText)
	vm.dispatchMemberSet(numberField.Num, "Value", updatedNumber)
	vm.adodbRecordsetUpdate(rs, nil)
	if vm.lastError != nil {
		t.Fatalf("update failed: %v", vm.lastError)
	}

	verifyText, verifyNumber := loadRow()
	if verifyText.String() != updatedText.String() {
		t.Fatalf("expected persisted text %q, got %q", updatedText.String(), verifyText.String())
	}
	if int64(vm.asInt(verifyNumber)) != updatedNumber.Num {
		t.Fatalf("expected persisted number %d, got %#v", updatedNumber.Num, verifyNumber)
	}
}

func TestADODBAccessRecordsetUpdatePersistsDateField(t *testing.T) {
	repoRoot, err := filepath.Abs("..")
	if err != nil {
		t.Fatalf("resolve repo root: %v", err)
	}
	dbPath := filepath.Join(repoRoot, "www", "asplite-test", "db", "sample.mdb")
	if _, err := os.Stat(dbPath); err != nil {
		t.Skipf("sample database not available: %v", err)
	}

	vm := NewVM(nil, nil, 5)
	defer vm.CleanupRequestResources()

	connVal := vm.newADODBConnection()
	conn := vm.adodbConnectionItems[connVal.Num]
	if conn == nil {
		t.Fatal("expected ADODB connection instance")
	}
	conn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath
	vm.adodbConnectionOpen(conn)
	if conn.state != adStateOpen {
		t.Fatalf("expected Access connection open, state=%d", conn.state)
	}

	firstRowVal := vm.adodbConnectionExecute(conn, []Value{NewString("select top 1 * from contact order by iId")})
	firstRow := vm.adodbRecordsetItems[firstRowVal.Num]
	if firstRow == nil || firstRow.recordCount == 0 {
		t.Fatal("expected at least one contact row in sample database")
	}
	targetID := vm.asInt(vm.dispatchMemberGet(firstRowVal, "iId"))
	if targetID == 0 {
		t.Fatalf("expected non-zero contact iId, got %d", targetID)
	}

	loadDate := func() Value {
		rsVal := vm.newADODBRecordset()
		rs := vm.adodbRecordsetItems[rsVal.Num]
		if rs == nil {
			t.Fatal("expected ADODB recordset instance")
		}
		rs.activeConnection = connVal.Num
		rs.cursorType = adOpenKeyset
		rs.lockType = adLockOptimistic
		sqlText := "select * from contact where iId=" + strconv.Itoa(targetID)
		vm.adodbRecordsetOpen(rs, sqlText, conn, []Value{NewString(sqlText)})
		if rs.state != adStateOpen || rs.recordCount != 1 {
			t.Fatalf("expected one contact row, state=%d count=%d", rs.state, rs.recordCount)
		}
		return vm.dispatchMemberGet(rsVal, "dDate")
	}

	originalDate := loadDate()
	defer func() {
		literal, ok := vm.adodbSQLLiteral(conn, originalDate)
		if !ok {
			return
		}
		restoreSQL := "UPDATE [contact] SET [dDate] = " + literal + " WHERE iId=" + strconv.Itoa(targetID)
		_, _ = vm.adodbExecWriteback(conn, restoreSQL, "TestADODBAccessRecordsetUpdatePersistsDateField", false)
	}()

	rsVal := vm.newADODBRecordset()
	rs := vm.adodbRecordsetItems[rsVal.Num]
	if rs == nil {
		t.Fatal("expected ADODB recordset instance")
	}
	rs.activeConnection = connVal.Num
	rs.cursorType = adOpenKeyset
	rs.lockType = adLockOptimistic
	sqlText := "select * from contact where iId=" + strconv.Itoa(targetID)
	vm.adodbRecordsetOpen(rs, sqlText, conn, []Value{NewString(sqlText)})
	if rs.state != adStateOpen || rs.recordCount != 1 {
		t.Fatalf("expected one contact row, state=%d count=%d", rs.state, rs.recordCount)
	}

	updatedDateText := "10/07/1974"
	dateField := vm.newADODBFieldProxy(rs, "dDate")
	vm.dispatchMemberSet(dateField.Num, "Value", NewString(updatedDateText))
	vm.adodbRecordsetUpdate(rs, nil)
	if vm.lastError != nil {
		t.Fatalf("update failed: %v", vm.lastError)
	}

	persistedDate := loadDate()
	persistedTime, ok := parseCompatDateValue(persistedDate.String())
	if !ok {
		t.Fatalf("expected persisted date parseable, got %#v", persistedDate)
	}
	if persistedTime.Year() != 1974 || persistedTime.Month() != time.July || persistedTime.Day() != 10 {
		t.Fatalf("expected persisted date 1974-07-10, got %s", persistedTime.Format(time.RFC3339))
	}
}

func TestADODBAccessRecordsetUpdateZeroDateStoresNull(t *testing.T) {
	repoRoot, err := filepath.Abs("..")
	if err != nil {
		t.Fatalf("resolve repo root: %v", err)
	}
	dbPath := filepath.Join(repoRoot, "www", "asplite-test", "db", "sample.mdb")
	if _, err := os.Stat(dbPath); err != nil {
		t.Skipf("sample database not available: %v", err)
	}

	vm := NewVM(nil, nil, 5)
	defer vm.CleanupRequestResources()

	connVal := vm.newADODBConnection()
	conn := vm.adodbConnectionItems[connVal.Num]
	if conn == nil {
		t.Fatal("expected ADODB connection instance")
	}
	conn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath
	vm.adodbConnectionOpen(conn)
	if conn.state != adStateOpen {
		t.Fatalf("expected Access connection open, state=%d", conn.state)
	}

	firstRowVal := vm.adodbConnectionExecute(conn, []Value{NewString("select top 1 * from contact order by iId")})
	firstRow := vm.adodbRecordsetItems[firstRowVal.Num]
	if firstRow == nil || firstRow.recordCount == 0 {
		t.Fatal("expected at least one contact row in sample database")
	}
	targetID := vm.asInt(vm.dispatchMemberGet(firstRowVal, "iId"))
	if targetID == 0 {
		t.Fatalf("expected non-zero contact iId, got %d", targetID)
	}

	loadDate := func() Value {
		rsVal := vm.newADODBRecordset()
		rs := vm.adodbRecordsetItems[rsVal.Num]
		if rs == nil {
			t.Fatal("expected ADODB recordset instance")
		}
		rs.activeConnection = connVal.Num
		rs.cursorType = adOpenKeyset
		rs.lockType = adLockOptimistic
		sqlText := "select * from contact where iId=" + strconv.Itoa(targetID)
		vm.adodbRecordsetOpen(rs, sqlText, conn, []Value{NewString(sqlText)})
		if rs.state != adStateOpen || rs.recordCount != 1 {
			t.Fatalf("expected one contact row, state=%d count=%d", rs.state, rs.recordCount)
		}
		return vm.dispatchMemberGet(rsVal, "dDate")
	}

	originalDate := loadDate()
	defer func() {
		literal, ok := vm.adodbSQLLiteral(conn, originalDate)
		if !ok {
			return
		}
		restoreSQL := "UPDATE [contact] SET [dDate] = " + literal + " WHERE iId=" + strconv.Itoa(targetID)
		_, _ = vm.adodbExecWriteback(conn, restoreSQL, "TestADODBAccessRecordsetUpdateZeroDateStoresNull", false)
	}()

	rsVal := vm.newADODBRecordset()
	rs := vm.adodbRecordsetItems[rsVal.Num]
	if rs == nil {
		t.Fatal("expected ADODB recordset instance")
	}
	rs.activeConnection = connVal.Num
	rs.cursorType = adOpenKeyset
	rs.lockType = adLockOptimistic
	sqlText := "select * from contact where iId=" + strconv.Itoa(targetID)
	vm.adodbRecordsetOpen(rs, sqlText, conn, []Value{NewString(sqlText)})
	if rs.state != adStateOpen || rs.recordCount != 1 {
		t.Fatalf("expected one contact row, state=%d count=%d", rs.state, rs.recordCount)
	}

	dateField := vm.newADODBFieldProxy(rs, "dDate")
	vm.dispatchMemberSet(dateField.Num, "Value", NewDate(time.Time{}))
	vm.adodbRecordsetUpdate(rs, nil)
	if vm.lastError != nil {
		t.Fatalf("update failed: %v", vm.lastError)
	}

	persistedDate := loadDate()
	if persistedDate.Type != VTNull && strings.TrimSpace(persistedDate.String()) != "" {
		parsed, ok := parseCompatDateValue(persistedDate.String())
		if ok && parsed.Year() == 1753 {
			t.Fatalf("unexpected SQL Server fallback-like date persisted: %s", parsed.Format(time.RFC3339))
		}
	}
}

func TestADODBAccessRecordsetAddNewDefaultMemberAssignment(t *testing.T) {
	repoRoot, err := filepath.Abs("..")
	if err != nil {
		t.Fatalf("resolve repo root: %v", err)
	}
	dbPath := filepath.Join(repoRoot, "www", "asplite-test", "ebook", "access.mdb")
	if _, err := os.Stat(dbPath); err != nil {
		t.Skipf("ebook database not available: %v", err)
	}

	vm := NewVM(nil, nil, 5)
	defer vm.CleanupRequestResources()

	connVal := vm.newADODBConnection()
	conn := vm.adodbConnectionItems[connVal.Num]
	if conn == nil {
		t.Fatal("expected ADODB connection instance")
	}
	conn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath
	vm.adodbConnectionOpen(conn)
	if conn.state != adStateOpen {
		t.Fatalf("expected Access connection open, state=%d", conn.state)
	}

	marker := strconv.FormatInt(time.Now().UnixNano(), 10)
	lastName := "AxonAddNewL_" + marker
	firstName := "AxonAddNewF_" + marker
	escapedLastName := strings.ReplaceAll(lastName, "'", "''")
	escapedFirstName := strings.ReplaceAll(firstName, "'", "''")
	cleanupSQL := "DELETE FROM [person] WHERE [lastname]='" + escapedLastName + "' AND [firstname]='" + escapedFirstName + "'"
	defer func() {
		_, _ = vm.adodbExecWriteback(conn, cleanupSQL, "TestADODBAccessRecordsetAddNewDefaultMemberAssignment", false)
	}()

	rsVal := vm.newADODBRecordset()
	rs := vm.adodbRecordsetItems[rsVal.Num]
	if rs == nil {
		t.Fatal("expected ADODB recordset instance")
	}
	rs.activeConnection = connVal.Num
	rs.cursorType = adOpenKeyset
	rs.lockType = adLockOptimistic
	sqlText := "select * from person where 1=2"
	vm.adodbRecordsetOpen(rs, sqlText, conn, []Value{NewString(sqlText)})
	if rs.state != adStateOpen {
		t.Fatalf("expected updatable recordset open, state=%d", rs.state)
	}

	vm.dispatchADODBRecordsetMethod(rs, "AddNew", nil)
	vm.dispatchADODBRecordsetMethod(rs, "", []Value{NewString("lastname"), NewString(lastName)})
	vm.dispatchADODBRecordsetMethod(rs, "", []Value{NewString("firstname"), NewString(firstName)})
	vm.dispatchADODBRecordsetMethod(rs, "Update", nil)
	if vm.lastError != nil {
		t.Fatalf("unexpected update error: %v", vm.lastError)
	}

	verifyRSVal := vm.adodbConnectionExecute(conn, []Value{NewString("select top 1 * from person where lastname='" + escapedLastName + "' and firstname='" + escapedFirstName + "'")})
	verifyRS := vm.adodbRecordsetItems[verifyRSVal.Num]
	if verifyRS == nil || verifyRS.recordCount == 0 {
		t.Fatalf("expected inserted row for lastname=%q firstname=%q", lastName, firstName)
	}
}

func parseCompatDateValue(text string) (time.Time, bool) {
	trimmed := strings.TrimSpace(text)
	if trimmed == "" {
		return time.Time{}, false
	}
	layouts := []string{
		"2006-01-02 15:04:05",
		"2006-01-02",
		time.RFC3339,
		"02/01/2006 15:04:05",
		"02/01/2006",
		"01/02/2006 15:04:05",
		"01/02/2006",
	}
	for i := 0; i < len(layouts); i++ {
		parsed, err := time.ParseInLocation(layouts[i], trimmed, time.Local)
		if err == nil {
			return parsed, true
		}
	}
	return time.Time{}, false
}
