/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
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
	"strings"
	"testing"
)

// TestVMServerFSOGetStandardStreamCompatibility verifies GetStandardStream returns TextStream objects with stable cursor properties.
func TestVMServerFSOGetStandardStreamCompatibility(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	host.Server().SetRequestPath("/tests/test_fso.asp")
	vm.SetHost(host)

	fso := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("Scripting.FileSystemObject")})
	stdout := vm.dispatchNativeCall(fso.Num, "GetStandardStream", []Value{NewInteger(1)})
	if stdout.Type != VTNativeObject {
		t.Fatalf("expected stdout TextStream object, got %#v", stdout)
	}
	if line := vm.dispatchMemberGet(stdout, "Line"); line.Type != VTInteger || line.Num != 1 {
		t.Fatalf("unexpected stdout Line property: %#v", line)
	}
	if column := vm.dispatchMemberGet(stdout, "Column"); column.Type != VTInteger || column.Num != 1 {
		t.Fatalf("unexpected stdout Column property: %#v", column)
	}

	vm.dispatchNativeCall(stdout.Num, "Close", nil)
	stderr := vm.dispatchNativeCall(fso.Num, "GetStandardStream", []Value{NewInteger(2)})
	if stderr.Type != VTNativeObject {
		t.Fatalf("expected stderr TextStream object after stdout close, got %#v", stderr)
	}
}

// TestVMADODBConnectionPropertiesAndOpenSchema verifies Connection compatibility properties and OpenSchema table enumeration.
func TestVMADODBConnectionPropertiesAndOpenSchema(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_adodb.asp")
	vm.SetHost(host)

	conn := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	if conn.Type != VTNativeObject {
		t.Fatalf("expected connection object, got %#v", conn)
	}

	dbPath := filepath.Join(rootDir, "schema.db")
	vm.dispatchMemberSet(conn.Num, "ConnectionString", NewString("sqlite:"+dbPath))
	vm.dispatchMemberSet(conn.Num, "CommandTimeout", NewInteger(12))
	vm.dispatchMemberSet(conn.Num, "ConnectionTimeout", NewInteger(7))
	vm.dispatchMemberSet(conn.Num, "CursorLocation", NewInteger(adUseClient))
	vm.dispatchMemberSet(conn.Num, "DefaultDatabase", NewString(dbPath))
	vm.dispatchMemberSet(conn.Num, "IsolationLevel", NewInteger(4096))
	vm.dispatchNativeCall(conn.Num, "Open", nil)
	defer vm.dispatchNativeCall(conn.Num, "Close", nil)
	vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("CREATE TABLE demo (id INTEGER PRIMARY KEY, name TEXT NOT NULL, amount NUMERIC(10,2))")})
	vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("CREATE TABLE parent (id INTEGER PRIMARY KEY)")})
	vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("CREATE TABLE child (id INTEGER PRIMARY KEY, parent_id INTEGER, CONSTRAINT fk_child_parent FOREIGN KEY(parent_id) REFERENCES parent(id))")})
	vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("CREATE UNIQUE INDEX idx_demo_name ON demo(name)")})
	vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("CREATE VIEW demo_view AS SELECT id, name FROM demo")})

	if provider := vm.dispatchMemberGet(conn, "Provider"); provider.Type != VTString || provider.Str != "SQLite" {
		t.Fatalf("unexpected Provider property: %#v", provider)
	}
	if version := vm.dispatchMemberGet(conn, "Version"); version.Type != VTString || strings.TrimSpace(version.Str) == "" {
		t.Fatalf("unexpected Version property: %#v", version)
	}
	if timeout := vm.dispatchMemberGet(conn, "CommandTimeout"); timeout.Type != VTInteger || timeout.Num != 12 {
		t.Fatalf("unexpected CommandTimeout property: %#v", timeout)
	}
	if timeout := vm.dispatchMemberGet(conn, "ConnectionTimeout"); timeout.Type != VTInteger || timeout.Num != 7 {
		t.Fatalf("unexpected ConnectionTimeout property: %#v", timeout)
	}
	if cursorLocation := vm.dispatchMemberGet(conn, "CursorLocation"); cursorLocation.Type != VTInteger || cursorLocation.Num != adUseClient {
		t.Fatalf("unexpected CursorLocation property: %#v", cursorLocation)
	}
	if defaultDatabase := vm.dispatchMemberGet(conn, "DefaultDatabase"); defaultDatabase.Type != VTString || defaultDatabase.Str != dbPath {
		t.Fatalf("unexpected DefaultDatabase property: %#v", defaultDatabase)
	}
	if isolation := vm.dispatchMemberGet(conn, "IsolationLevel"); isolation.Type != VTInteger || isolation.Num != 4096 {
		t.Fatalf("unexpected IsolationLevel property: %#v", isolation)
	}

	tableRestrictions := Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{NewEmpty(), NewEmpty(), NewString("demo"), NewString("TABLE")})}
	schema := vm.dispatchNativeCall(conn.Num, "OpenSchema", []Value{NewInteger(adSchemaTables), tableRestrictions})
	if schema.Type != VTNativeObject {
		t.Fatalf("expected schema recordset object, got %#v", schema)
	}
	if state := vm.dispatchMemberGet(schema, "State"); state.Type != VTInteger || state.Num != adStateOpen {
		t.Fatalf("unexpected schema State property: %#v", state)
	}
	if recordCount := vm.dispatchMemberGet(schema, "RecordCount"); recordCount.Type != VTInteger || recordCount.Num != 1 {
		t.Fatalf("unexpected schema RecordCount property: %#v", recordCount)
	}
	if tableName := vm.dispatchMemberGet(schema, "TABLE_NAME"); tableName.Type != VTString || !strings.EqualFold(tableName.Str, "demo") {
		t.Fatalf("unexpected schema table name: %#v", tableName)
	}

	columnRestrictions := Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{NewEmpty(), NewEmpty(), NewString("demo"), NewString("name")})}
	columnsSchema := vm.dispatchNativeCall(conn.Num, "OpenSchema", []Value{NewInteger(adSchemaColumns), columnRestrictions})
	if columnsSchema.Type != VTNativeObject {
		t.Fatalf("expected columns schema recordset object, got %#v", columnsSchema)
	}
	if recordCount := vm.dispatchMemberGet(columnsSchema, "RecordCount"); recordCount.Type != VTInteger || recordCount.Num != 1 {
		t.Fatalf("unexpected columns schema RecordCount: %#v", recordCount)
	}
	if tableName := vm.dispatchMemberGet(columnsSchema, "TABLE_NAME"); tableName.Type != VTString || !strings.EqualFold(tableName.Str, "demo") {
		t.Fatalf("unexpected columns schema table name: %#v", tableName)
	}
	if columnName := vm.dispatchMemberGet(columnsSchema, "COLUMN_NAME"); columnName.Type != VTString || !strings.EqualFold(columnName.Str, "name") {
		t.Fatalf("unexpected columns schema column name: %#v", columnName)
	}
	if dataType := vm.dispatchMemberGet(columnsSchema, "DATA_TYPE"); dataType.Type != VTString || !strings.EqualFold(dataType.Str, "TEXT") {
		t.Fatalf("unexpected columns schema data type: %#v", dataType)
	}
	if isNullable := vm.dispatchMemberGet(columnsSchema, "IS_NULLABLE"); isNullable.Type != VTString || !strings.EqualFold(isNullable.Str, "NO") {
		t.Fatalf("unexpected columns schema nullability: %#v", isNullable)
	}

	indexRestrictions := Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{NewEmpty(), NewEmpty(), NewString("idx_demo_name"), NewEmpty(), NewString("demo")})}
	indexesSchema := vm.dispatchNativeCall(conn.Num, "OpenSchema", []Value{NewInteger(adSchemaIndexes), indexRestrictions})
	if indexesSchema.Type != VTNativeObject {
		t.Fatalf("expected indexes schema recordset object, got %#v", indexesSchema)
	}
	if recordCount := vm.dispatchMemberGet(indexesSchema, "RecordCount"); recordCount.Type != VTInteger || recordCount.Num != 1 {
		t.Fatalf("unexpected indexes schema RecordCount: %#v", recordCount)
	}
	if indexName := vm.dispatchMemberGet(indexesSchema, "INDEX_NAME"); indexName.Type != VTString || !strings.EqualFold(indexName.Str, "idx_demo_name") {
		t.Fatalf("unexpected indexes schema index name: %#v", indexName)
	}
	if tableName := vm.dispatchMemberGet(indexesSchema, "TABLE_NAME"); tableName.Type != VTString || !strings.EqualFold(tableName.Str, "demo") {
		t.Fatalf("unexpected indexes schema table name: %#v", tableName)
	}
	if uniqueValue := vm.dispatchMemberGet(indexesSchema, "UNIQUE"); uniqueValue.Type != VTBool || uniqueValue.Num != 1 {
		t.Fatalf("unexpected indexes schema UNIQUE flag: %#v", uniqueValue)
	}
	if columnName := vm.dispatchMemberGet(indexesSchema, "COLUMN_NAME"); columnName.Type != VTString || !strings.EqualFold(columnName.Str, "name") {
		t.Fatalf("unexpected indexes schema column name: %#v", columnName)
	}

	viewRestrictions := Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{NewEmpty(), NewEmpty(), NewString("demo_view")})}
	viewsSchema := vm.dispatchNativeCall(conn.Num, "OpenSchema", []Value{NewInteger(adSchemaViews), viewRestrictions})
	if viewsSchema.Type != VTNativeObject {
		t.Fatalf("expected views schema recordset object, got %#v", viewsSchema)
	}
	if recordCount := vm.dispatchMemberGet(viewsSchema, "RecordCount"); recordCount.Type != VTInteger || recordCount.Num != 1 {
		t.Fatalf("unexpected views schema RecordCount: %#v", recordCount)
	}
	if tableName := vm.dispatchMemberGet(viewsSchema, "TABLE_NAME"); tableName.Type != VTString || !strings.EqualFold(tableName.Str, "demo_view") {
		t.Fatalf("unexpected views schema table name: %#v", tableName)
	}

	foreignKeyRestrictions := Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{NewEmpty(), NewEmpty(), NewString("parent"), NewEmpty(), NewEmpty(), NewString("child")})}
	foreignKeysSchema := vm.dispatchNativeCall(conn.Num, "OpenSchema", []Value{NewInteger(adSchemaForeignKeys), foreignKeyRestrictions})
	if foreignKeysSchema.Type != VTNativeObject {
		t.Fatalf("expected foreign keys schema recordset object, got %#v", foreignKeysSchema)
	}
	if recordCount := vm.dispatchMemberGet(foreignKeysSchema, "RecordCount"); recordCount.Type != VTInteger || recordCount.Num != 1 {
		t.Fatalf("unexpected foreign keys schema RecordCount: %#v", recordCount)
	}
	if pkTableName := vm.dispatchMemberGet(foreignKeysSchema, "PK_TABLE_NAME"); pkTableName.Type != VTString || !strings.EqualFold(pkTableName.Str, "parent") {
		t.Fatalf("unexpected foreign keys schema PK table name: %#v", pkTableName)
	}
	if fkTableName := vm.dispatchMemberGet(foreignKeysSchema, "FK_TABLE_NAME"); fkTableName.Type != VTString || !strings.EqualFold(fkTableName.Str, "child") {
		t.Fatalf("unexpected foreign keys schema FK table name: %#v", fkTableName)
	}
	if fkColumnName := vm.dispatchMemberGet(foreignKeysSchema, "FK_COLUMN_NAME"); fkColumnName.Type != VTString || !strings.EqualFold(fkColumnName.Str, "parent_id") {
		t.Fatalf("unexpected foreign keys schema FK column name: %#v", fkColumnName)
	}
	if keySeq := vm.dispatchMemberGet(foreignKeysSchema, "KEY_SEQ"); keySeq.Type != VTInteger || keySeq.Num != 1 {
		t.Fatalf("unexpected foreign keys schema KEY_SEQ: %#v", keySeq)
	}

	proceduresSchema := vm.dispatchNativeCall(conn.Num, "OpenSchema", []Value{NewInteger(adSchemaProcedures)})
	if proceduresSchema.Type != VTNativeObject {
		t.Fatalf("expected procedures schema recordset object, got %#v", proceduresSchema)
	}
	if recordCount := vm.dispatchMemberGet(proceduresSchema, "RecordCount"); recordCount.Type != VTInteger || recordCount.Num != 0 {
		t.Fatalf("unexpected procedures schema RecordCount: %#v", recordCount)
	}
}

// TestVMADODBSQLiteDataSourceExecuteParams verifies SQLite Data Source parsing,
// ADODB.Connection.Execute parameter arrays, and transaction-aware Execute behavior.
func TestVMADODBSQLiteDataSourceExecuteParams(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_adodb_sqlite_execute.asp")
	vm.SetHost(host)

	conn := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	if conn.Type != VTNativeObject {
		t.Fatalf("expected connection object, got %#v", conn)
	}

	dbPath := filepath.Join(rootDir, "imported_data.db")
	connStr := "Driver={SQLite3};Data Source=" + dbPath
	vm.dispatchMemberSet(conn.Num, "ConnectionString", NewString(connStr))
	vm.dispatchNativeCall(conn.Num, "Open", nil)
	defer vm.dispatchNativeCall(conn.Num, "Close", nil)

	if provider := vm.dispatchMemberGet(conn, "Provider"); provider.Type != VTString || provider.Str != "SQLite" {
		t.Fatalf("unexpected Provider property: %#v", provider)
	}
	if defaultDatabase := vm.dispatchMemberGet(conn, "DefaultDatabase"); defaultDatabase.Type != VTString || defaultDatabase.Str != dbPath {
		t.Fatalf("unexpected DefaultDatabase property: %#v", defaultDatabase)
	}

	vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("CREATE TABLE [tblLabel] ([Id] INTEGER, [Name] TEXT)")})
	vm.dispatchNativeCall(conn.Num, "BeginTrans", nil)

	insertParams := Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{NewInteger(1), NewString("Label One")})}
	vm.dispatchNativeCall(conn.Num, "Execute", []Value{
		NewString("INSERT INTO [tblLabel] ([Id], [Name]) VALUES (?, ?)"),
		insertParams,
	})
	vm.dispatchNativeCall(conn.Num, "CommitTrans", nil)

	countRS := vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("SELECT COUNT(*) AS C FROM [tblLabel]")})
	if countRS.Type != VTNativeObject {
		t.Fatalf("expected count recordset, got %#v", countRS)
	}
	count := vm.dispatchMemberGet(countRS, "C")
	if count.Type != VTInteger || count.Num != 1 {
		t.Fatalf("unexpected inserted row count: %#v", count)
	}

	if _, err := os.Stat(dbPath); err != nil {
		t.Fatalf("expected SQLite file to exist at %s: %v", dbPath, err)
	}
}

// TestVMADODBRecordsetCompatibilitySurface verifies newly implemented Recordset and Field compatibility members.
func TestVMADODBRecordsetCompatibilitySurface(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	host.Server().SetRequestPath("/tests/test_adodb_recordset.asp")
	vm.SetHost(host)

	rs := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Recordset")})
	cmd := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Command")})
	stream := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Stream")})
	if rs.Type != VTNativeObject || cmd.Type != VTNativeObject || stream.Type != VTNativeObject {
		t.Fatalf("expected ADODB native objects, got rs=%#v cmd=%#v stream=%#v", rs, cmd, stream)
	}

	vm.dispatchNativeCall(rs.Num, "Fields.Append", []Value{NewString("Id"), NewInteger(3), NewInteger(10)})
	vm.dispatchNativeCall(rs.Num, "Fields.Append", []Value{NewString("Name"), NewInteger(200), NewInteger(50)})
	vm.dispatchNativeCall(rs.Num, "Open", nil)
	vm.dispatchMemberSet(cmd.Num, "CommandText", NewString("SELECT * FROM demo"))
	vm.dispatchMemberSet(rs.Num, "ActiveCommand", cmd)
	vm.dispatchMemberSet(rs.Num, "DataMember", NewString("main"))
	vm.dispatchMemberSet(rs.Num, "Index", NewString("Id"))
	vm.dispatchMemberSet(rs.Num, "MarshalOptions", NewInteger(2))
	vm.dispatchMemberSet(rs.Num, "MaxRecords", NewInteger(25))

	vm.dispatchNativeCall(rs.Num, "AddNew", nil)
	idField := vm.dispatchNativeCall(rs.Num, "Fields.Item", []Value{NewString("Id")})
	nameField := vm.dispatchNativeCall(rs.Num, "Fields.Item", []Value{NewString("Name")})
	vm.dispatchMemberSet(idField.Num, "Value", NewInteger(10))
	vm.dispatchMemberSet(nameField.Num, "Value", NewString("Axon"))
	vm.dispatchNativeCall(rs.Num, "Update", nil)

	if activeCommand := vm.dispatchMemberGet(rs, "ActiveCommand"); activeCommand.Type != VTNativeObject || activeCommand.Num != cmd.Num {
		t.Fatalf("unexpected ActiveCommand property: %#v", activeCommand)
	}
	if dataMember := vm.dispatchMemberGet(rs, "DataMember"); dataMember.Type != VTString || dataMember.Str != "main" {
		t.Fatalf("unexpected DataMember property: %#v", dataMember)
	}
	if editMode := vm.dispatchMemberGet(rs, "EditMode"); editMode.Type != VTInteger || editMode.Num != adEditNone {
		t.Fatalf("unexpected EditMode property: %#v", editMode)
	}
	if index := vm.dispatchMemberGet(rs, "Index"); index.Type != VTString || index.Str != "Id" {
		t.Fatalf("unexpected Index property: %#v", index)
	}
	if marshalOptions := vm.dispatchMemberGet(rs, "MarshalOptions"); marshalOptions.Type != VTInteger || marshalOptions.Num != 2 {
		t.Fatalf("unexpected MarshalOptions property: %#v", marshalOptions)
	}
	if maxRecords := vm.dispatchMemberGet(rs, "MaxRecords"); maxRecords.Type != VTInteger || maxRecords.Num != 25 {
		t.Fatalf("unexpected MaxRecords property: %#v", maxRecords)
	}
	if source := vm.dispatchMemberGet(rs, "Source"); source.Type != VTString || source.Str != "SELECT * FROM demo" {
		t.Fatalf("unexpected Source property: %#v", source)
	}
	if status := vm.dispatchMemberGet(rs, "Status"); status.Type != VTInteger || status.Num != adRecOK {
		t.Fatalf("unexpected Status property: %#v", status)
	}

	field := nameField
	if field.Type != VTNativeObject {
		t.Fatalf("expected field proxy object, got %#v", field)
	}
	if actualSize := vm.dispatchMemberGet(field, "ActualSize"); actualSize.Type != VTInteger || actualSize.Num != 4 {
		t.Fatalf("unexpected ActualSize property: %#v", actualSize)
	}
	if precision := vm.dispatchMemberGet(field, "Precision"); precision.Type != VTInteger || precision.Num != 50 {
		t.Fatalf("unexpected Precision property: %#v", precision)
	}
	if original := vm.dispatchMemberGet(field, "OriginalValue"); original.Type != VTString || original.Str != "Axon" {
		t.Fatalf("unexpected OriginalValue property: %#v", original)
	}
	if underlying := vm.dispatchMemberGet(field, "UnderlyingValue"); underlying.Type != VTString || underlying.Str != "Axon" {
		t.Fatalf("unexpected UnderlyingValue property: %#v", underlying)
	}
	if fieldStatus := vm.dispatchMemberGet(field, "Status"); fieldStatus.Type != VTInteger || fieldStatus.Num != adRecOK {
		t.Fatalf("unexpected field Status property: %#v", fieldStatus)
	}

	if compare := vm.dispatchNativeCall(rs.Num, "CompareBookmarks", []Value{NewInteger(1), NewInteger(1)}); compare.Type != VTInteger || compare.Num != 0 {
		t.Fatalf("unexpected CompareBookmarks result: %#v", compare)
	}
	vm.dispatchNativeCall(rs.Num, "Seek", []Value{NewString("10")})
	if position := vm.dispatchMemberGet(rs, "AbsolutePosition"); position.Type != VTInteger || position.Num != 1 {
		t.Fatalf("unexpected AbsolutePosition after Seek: %#v", position)
	}
	if supports := vm.dispatchNativeCall(rs.Num, "Supports", []Value{NewInteger(1)}); supports.Type != VTBool || supports.Num != 1 {
		t.Fatalf("unexpected Supports result: %#v", supports)
	}

	vm.dispatchNativeCall(stream.Num, "Open", nil)
	vm.dispatchNativeCall(rs.Num, "Save", []Value{stream})
	if size := vm.dispatchMemberGet(stream, "Size"); size.Type != VTInteger || size.Num <= 0 {
		t.Fatalf("unexpected stream Size after Recordset.Save: %#v", size)
	}

	clone := vm.dispatchNativeCall(rs.Num, "Clone", nil)
	if clone.Type != VTNativeObject {
		t.Fatalf("expected Clone to return native object, got %#v", clone)
	}
	vm.dispatchNativeCall(rs.Num, "CancelUpdate", nil)
	vm.dispatchNativeCall(rs.Num, "UpdateBatch", nil)
	vm.dispatchNativeCall(rs.Num, "Requery", nil)
	vm.dispatchNativeCall(rs.Num, "Resync", nil)
	if recordCount := vm.dispatchMemberGet(rs, "RecordCount"); recordCount.Type != VTInteger || recordCount.Num != 1 {
		t.Fatalf("unexpected RecordCount after Requery/Resync: %#v", recordCount)
	}
}

// TestVMADODBRecordsetUpdateOnlyDirtyFields verifies disconnected writeback updates only changed columns.
func TestVMADODBRecordsetUpdateOnlyDirtyFields(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_adodb_recordset_dirty_update.asp")
	vm.SetHost(host)

	conn := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	if conn.Type != VTNativeObject {
		t.Fatalf("expected connection object, got %#v", conn)
	}

	dbPath := filepath.Join(rootDir, "dirty_update.db")
	vm.dispatchMemberSet(conn.Num, "ConnectionString", NewString("sqlite:"+dbPath))
	vm.dispatchNativeCall(conn.Num, "Open", nil)
	defer vm.dispatchNativeCall(conn.Num, "Close", nil)

	vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("CREATE TABLE demo (id INTEGER PRIMARY KEY, name TEXT NOT NULL, slug TEXT GENERATED ALWAYS AS (lower(name)) VIRTUAL)")})
	vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("INSERT INTO demo (id, name) VALUES (1, 'Alpha')")})

	rs := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Recordset")})
	if rs.Type != VTNativeObject {
		t.Fatalf("expected recordset object, got %#v", rs)
	}
	vm.dispatchMemberSet(rs.Num, "ActiveConnection", conn)
	vm.dispatchNativeCall(rs.Num, "Open", []Value{NewString("select * from demo where id=1")})

	vm.dispatchNativeCall(rs.Num, "", []Value{NewString("name"), NewString("Beta")})
	vm.dispatchNativeCall(rs.Num, "Update", nil)
	if vm.lastError != nil {
		t.Fatalf("recordset update failed: %v", vm.lastError)
	}

	verify := vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("SELECT name, slug FROM demo WHERE id=1")})
	if verify.Type != VTNativeObject {
		t.Fatalf("expected verify recordset object, got %#v", verify)
	}
	name := vm.dispatchMemberGet(verify, "name")
	if name.Type != VTString || name.Str != "Beta" {
		t.Fatalf("unexpected persisted name: %#v", name)
	}
	slug := vm.dispatchMemberGet(verify, "slug")
	if slug.Type != VTString || slug.Str != "beta" {
		t.Fatalf("unexpected generated slug value: %#v", slug)
	}
}

// TestVMADODBRecordsetNextRecordsetCompatibility verifies NextRecordset returns the next result set object when supported.
func TestVMADODBRecordsetNextRecordsetCompatibility(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_adodb_nextrecordset.asp")
	vm.SetHost(host)

	conn := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	if conn.Type != VTNativeObject {
		t.Fatalf("expected connection object, got %#v", conn)
	}
	dbPath := filepath.Join(rootDir, "nextrecordset.db")
	vm.dispatchMemberSet(conn.Num, "ConnectionString", NewString("sqlite:"+dbPath))
	vm.dispatchNativeCall(conn.Num, "Open", nil)
	defer vm.dispatchNativeCall(conn.Num, "Close", nil)

	rs := vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("SELECT 1 AS id; SELECT 2 AS id")})
	if rs.Type != VTNativeObject {
		t.Fatalf("expected first recordset object, got %#v", rs)
	}

	firstID := vm.dispatchMemberGet(rs, "id")
	if firstID.Type != VTInteger || (firstID.Num != 1 && firstID.Num != 2) {
		t.Fatalf("unexpected first resultset row value: %#v", firstID)
	}

	next := vm.dispatchNativeCall(rs.Num, "NextRecordset", nil)
	if next.Type == VTNativeObject {
		if next.Num == rs.Num {
			t.Fatalf("expected NextRecordset to allocate a distinct recordset object")
		}
		if bof := vm.dispatchMemberGet(next, "BOF"); bof.Type != VTBool || bof.Num != 0 {
			t.Fatalf("unexpected BOF on second resultset: %#v", bof)
		}
		if eof := vm.dispatchMemberGet(next, "EOF"); eof.Type != VTBool || eof.Num != 0 {
			t.Fatalf("unexpected EOF on second resultset: %#v", eof)
		}
		secondID := vm.dispatchMemberGet(next, "id")
		if secondID.Type != VTInteger || secondID.Num == firstID.Num {
			t.Fatalf("unexpected second resultset row value: %#v", secondID)
		}
		end := vm.dispatchNativeCall(next.Num, "NextRecordset", nil)
		if end.Type != VTObject || end.Num != 0 {
			t.Fatalf("expected Nothing after final resultset, got %#v", end)
		}
		return
	}

	if next.Type != VTObject || next.Num != 0 {
		t.Fatalf("expected Nothing or native object from NextRecordset, got %#v", next)
	}
}

// TestVMADODBTypeNameCompatibility verifies that TypeName() returns Classic ASP-compatible strings
// for all ADODB native objects so that ASP code like asplite.asp can branch on the type correctly.
func TestVMADODBTypeNameCompatibility(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_typename.asp")
	vm.SetHost(host)

	typeNameOf := func(val Value) string {
		name, _ := vbsTypeNameVM(vm, []Value{val})
		return name.String()
	}

	// ADODB.Connection
	conn := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	if conn.Type != VTNativeObject {
		t.Fatalf("expected ADODB.Connection object, got %#v", conn)
	}
	if name := typeNameOf(conn); name != "Connection" {
		t.Errorf("TypeName(ADODB.Connection) = %q, want %q", name, "Connection")
	}

	// ADODB.Recordset (disconnected)
	rs := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Recordset")})
	if rs.Type != VTNativeObject {
		t.Fatalf("expected ADODB.Recordset object, got %#v", rs)
	}
	if name := typeNameOf(rs); name != "Recordset" {
		t.Errorf("TypeName(ADODB.Recordset) = %q, want %q", name, "Recordset")
	}

	// ADODB.Command
	cmd := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Command")})
	if cmd.Type != VTNativeObject {
		t.Fatalf("expected ADODB.Command object, got %#v", cmd)
	}
	if name := typeNameOf(cmd); name != "Command" {
		t.Errorf("TypeName(ADODB.Command) = %q, want %q", name, "Command")
	}

	// Fields collection
	vm.dispatchNativeCall(rs.Num, "Fields.Append", []Value{NewString("Title"), NewInteger(200), NewInteger(50)})
	vm.dispatchNativeCall(rs.Num, "Open", nil)
	fields := vm.dispatchMemberGet(rs, "Fields")
	if fields.Type != VTNativeObject {
		t.Fatalf("expected Fields collection object, got %#v", fields)
	}
	if name := typeNameOf(fields); name != "Fields" {
		t.Errorf("TypeName(Fields) = %q, want %q", name, "Fields")
	}

	// Single Field proxy
	field := vm.dispatchNativeCall(fields.Num, "Item", []Value{NewInteger(0)})
	if field.Type != VTNativeObject {
		t.Fatalf("expected Field object, got %#v", field)
	}
	if name := typeNameOf(field); name != "Field" {
		t.Errorf("TypeName(Field) = %q, want %q", name, "Field")
	}

	// ADODB.Stream
	stream := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Stream")})
	if stream.Type != VTNativeObject {
		t.Fatalf("expected ADODB.Stream object, got %#v", stream)
	}
	if name := typeNameOf(stream); name != "Stream" {
		t.Errorf("TypeName(ADODB.Stream) = %q, want %q", name, "Stream")
	}
}
