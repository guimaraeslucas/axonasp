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

// Package axonvm - lib_g3db.go
//
// G3DB is a lightweight native database library for AxonASP that exposes
// Go's standard database/sql interface directly to VBScript running inside
// the AxonASP VM. It supports MySQL, PostgreSQL, Microsoft SQL Server,
// SQLite, and Oracle (via go-ora/v2) without requiring any external COM
// components or ADODB.
//
// VBScript Usage:
//   Set db = Server.CreateObject("G3DB")
//   db.Open "mysql", "user:pass@tcp(localhost:3306)/dbname"
//   Set rs = db.Query("SELECT id, name FROM users WHERE active = ?", 1)
//   Do While Not rs.EOF
//     Response.Write rs("name") & "<br>"
//     rs.MoveNext
//   Loop
//   rs.Close
//   db.Close

import (
	"context"
	"database/sql"
	"fmt"
	"strings"
	"sync"
	"time"

	"g3pix.com.br/axonasp/axonconfig"
	_ "github.com/denisenkom/go-mssqldb"
	_ "github.com/go-sql-driver/mysql"
	_ "github.com/lib/pq"
	_ "github.com/sijms/go-ora/v2"
	"github.com/spf13/viper"
	_ "modernc.org/sqlite"
)

// ---------------------------------------------------------------------------
// G3DB — Primary Connection Object
// ---------------------------------------------------------------------------

// G3DB wraps a *sql.DB connection for VBScript consumption.
// All fields accessed concurrently are protected by mu.
type G3DB struct {
	vm        *VM
	db        *sql.DB
	driver    string
	dsn       string
	isOpen    bool
	lastError string
	mu        sync.RWMutex
}

// newG3DBObject registers a fresh G3DB instance in the VM and returns its handle Value.
func (vm *VM) newG3DBObject() Value {
	obj := &G3DB{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3dbItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet resolves read property access for G3DB (e.g. db.IsOpen, db.Driver).
func (g *G3DB) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(propertyName) {
	case "isopen":
		g.mu.RLock()
		v := g.isOpen
		g.mu.RUnlock()
		return NewBool(v)
	case "driver":
		g.mu.RLock()
		v := g.driver
		g.mu.RUnlock()
		return NewString(v)
	case "dsn":
		g.mu.RLock()
		v := g.dsn
		g.mu.RUnlock()
		return NewString(v)
	case "lasterror":
		g.mu.RLock()
		v := g.lastError
		g.mu.RUnlock()
		return NewString(v)
	}
	return g.DispatchMethod(propertyName, nil)
}

// DispatchPropertySet handles writable property assignments (e.g. db.Driver = "mysql").
func (g *G3DB) DispatchPropertySet(propertyName string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	val := args[0].String()
	switch strings.ToLower(propertyName) {
	case "driver":
		g.mu.Lock()
		g.driver = g3dbNormalizeDriver(val)
		g.mu.Unlock()
		return true
	case "dsn":
		g.mu.Lock()
		g.dsn = val
		g.mu.Unlock()
		return true
	}
	return false
}

// DispatchMethod routes all VBScript method calls for the G3DB connection object
// using a fast O(1) switch on the lower-cased method name.
func (g *G3DB) DispatchMethod(methodName string, args []Value) Value {
	switch strings.ToLower(methodName) {

	// Open(driver, dsn) — opens a connection to the specified database.
	case "open":
		if len(args) < 2 {
			g.setError(ErrG3DBRequiresDriverAndDSN.String())
			return NewBool(false)
		}
		return NewBool(g.open(args[0].String(), args[1].String()))

	// OpenFromEnv([driver]) — opens a connection using settings from axonasp.toml / ENV.
	case "openfromenv":
		driver := "mysql"
		if len(args) > 0 && args[0].String() != "" {
			driver = args[0].String()
		}
		return NewBool(g.openFromEnv(driver))

	// Close() — closes the database connection and releases the underlying *sql.DB.
	case "close":
		return NewBool(g.close())

	// Query(sql [, params...]) — executes a SELECT and returns a G3DBResultSet.
	case "query":
		if len(args) < 1 {
			g.setError(ErrG3DBQueryRequiresSQL.String())
			return NewEmpty()
		}
		return g.query(args[0].String(), args[1:])

	// QueryRow(sql [, params...]) — executes a SELECT expected to return one row.
	case "queryrow":
		if len(args) < 1 {
			g.setError(ErrG3DBQueryRequiresSQL.String())
			return NewEmpty()
		}
		return g.queryRow(args[0].String(), args[1:])

	// Exec(sql [, params...]) — executes an INSERT/UPDATE/DELETE statement.
	case "exec":
		if len(args) < 1 {
			g.setError(ErrG3DBExecRequiresSQL.String())
			return NewEmpty()
		}
		return g.exec(args[0].String(), args[1:])

	// Prepare(sql) — creates a prepared statement.
	case "prepare":
		if len(args) < 1 {
			g.setError(ErrG3DBPrepareRequiresSQL.String())
			return NewEmpty()
		}
		return g.prepare(args[0].String())

	// Begin / BeginTrans / BeginTransaction — starts a simple transaction.
	case "begin", "begintrans", "begintransaction":
		return g.beginTx(0, false)

	// BeginTx([timeoutSeconds, readOnly]) — starts a transaction with optional options.
	case "begintx":
		timeout := 0
		readOnly := false
		if len(args) > 0 {
			timeout = g.vm.asInt(args[0])
		}
		if len(args) > 1 {
			readOnly = args[1].Type == VTBool && args[1].Num != 0
		}
		return g.beginTx(timeout, readOnly)

	// SetMaxOpenConns(n) — limits the total number of open connections in the pool.
	case "setmaxopenconns":
		if len(args) > 0 {
			g.setMaxOpenConns(g.vm.asInt(args[0]))
		}
		return NewEmpty()

	// SetMaxIdleConns(n) — limits the number of idle connections retained in the pool.
	case "setmaxidleconns":
		if len(args) > 0 {
			g.setMaxIdleConns(g.vm.asInt(args[0]))
		}
		return NewEmpty()

	// SetConnMaxLifetime(seconds) — sets the maximum time a connection may be reused.
	case "setconnmaxlifetime":
		if len(args) > 0 {
			g.setConnMaxLifetime(g.vm.asInt(args[0]))
		}
		return NewEmpty()

	// SetConnMaxIdleTime(seconds) — sets the maximum idle time before a connection is closed.
	case "setconnmaxidletime":
		if len(args) > 0 {
			g.setConnMaxIdleTime(g.vm.asInt(args[0]))
		}
		return NewEmpty()

	// Stats() — returns a Scripting.Dictionary with runtime connection-pool statistics.
	case "stats":
		return g.stats()

	// GetLastError / GetError — returns the most recent error message string.
	case "geterror", "getlasterror":
		g.mu.RLock()
		e := g.lastError
		g.mu.RUnlock()
		return NewString(e)
	}
	return NewEmpty()
}

// open establishes a database connection using the provided driver and DSN.
// A 5-second ping is used to verify the connection before accepting it.
func (g *G3DB) open(driver, dsn string) bool {
	g.mu.Lock()
	defer g.mu.Unlock()

	if g.isOpen {
		g.lastError = ErrG3DBConnectionAlreadyOpen.String()
		return false
	}

	driver = g3dbNormalizeDriver(driver)
	if driver == "" {
		g.lastError = ErrG3DBUnsupportedDriver.String()
		return false
	}

	db, err := sql.Open(driver, dsn)
	if err != nil {
		g.lastError = ErrG3DBPingFailed.String() + ": " + err.Error()
		return false
	}

	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	if err := db.PingContext(ctx); err != nil {
		_ = db.Close()
		g.lastError = ErrG3DBPingFailed.String() + ": " + err.Error()
		return false
	}

	g.db = db
	g.driver = driver
	g.dsn = dsn
	g.isOpen = true
	g.lastError = ""
	return true
}

// openFromEnv reads connection parameters from axonasp.toml (with ENV override)
// and opens a connection to the specified driver.
func (g *G3DB) openFromEnv(driver string) bool {
	driver = g3dbNormalizeDriver(driver)
	if driver == "" {
		g.setError(ErrG3DBUnsupportedDriver.String())
		return false
	}

	v := g3dbNewConfigViper()
	dsn := g3dbBuildDSN(v, driver)
	if dsn == "" {
		g.setError(ErrG3DBMissingConfigKeys.String())
		return false
	}

	return g.open(driver, dsn)
}

// close shuts down the underlying *sql.DB connection pool.
func (g *G3DB) close() bool {
	g.mu.Lock()
	defer g.mu.Unlock()

	if !g.isOpen || g.db == nil {
		return true
	}
	err := g.db.Close()
	g.db = nil
	g.isOpen = false
	if err != nil {
		g.lastError = err.Error()
		return false
	}
	g.lastError = ""
	return true
}

// query executes a SELECT statement and returns a G3DBResultSet Value.
func (g *G3DB) query(sqlText string, params []Value) Value {
	g.mu.RLock()
	db := g.db
	driver := g.driver
	g.mu.RUnlock()

	if db == nil {
		g.setError(ErrG3DBConnectionNotOpen.String())
		return NewEmpty()
	}

	prepared := g3dbRewritePlaceholders(sqlText, driver)
	iArgs := g3dbValueSliceToInterface(params)
	rows, err := db.Query(prepared, iArgs...)
	if err != nil {
		g.setError(ErrG3DBQueryFailed.String() + ": " + err.Error())
		return NewEmpty()
	}

	g.setError("")
	return g.vm.newG3DBResultSetValue(rows)
}

// queryRow executes a SELECT expected to return at most one row.
func (g *G3DB) queryRow(sqlText string, params []Value) Value {
	g.mu.RLock()
	db := g.db
	driver := g.driver
	g.mu.RUnlock()

	if db == nil {
		g.setError(ErrG3DBConnectionNotOpen.String())
		return NewEmpty()
	}

	prepared := g3dbRewritePlaceholders(sqlText, driver)
	iArgs := g3dbValueSliceToInterface(params)
	row := db.QueryRow(prepared, iArgs...)
	g.setError("")
	return g.vm.newG3DBRowValue(row)
}

// exec runs an INSERT, UPDATE, or DELETE statement.
func (g *G3DB) exec(sqlText string, params []Value) Value {
	g.mu.RLock()
	db := g.db
	driver := g.driver
	g.mu.RUnlock()

	if db == nil {
		g.setError(ErrG3DBConnectionNotOpen.String())
		return NewEmpty()
	}

	prepared := g3dbRewritePlaceholders(sqlText, driver)
	iArgs := g3dbValueSliceToInterface(params)
	result, err := db.Exec(prepared, iArgs...)
	if err != nil {
		g.setError(ErrG3DBExecFailed.String() + ": " + err.Error())
		return NewEmpty()
	}

	g.setError("")
	return g.vm.newG3DBResultValue(result)
}

// prepare creates a reusable prepared statement.
func (g *G3DB) prepare(sqlText string) Value {
	g.mu.RLock()
	db := g.db
	driver := g.driver
	g.mu.RUnlock()

	if db == nil {
		g.setError(ErrG3DBConnectionNotOpen.String())
		return NewEmpty()
	}

	prepared := g3dbRewritePlaceholders(sqlText, driver)
	stmt, err := db.Prepare(prepared)
	if err != nil {
		g.setError(ErrG3DBPrepareFailed.String() + ": " + err.Error())
		return NewEmpty()
	}

	g.setError("")
	return g.vm.newG3DBStatementValue(stmt, driver)
}

// beginTx starts a database transaction with optional timeout and read-only flag.
func (g *G3DB) beginTx(timeoutSeconds int, readOnly bool) Value {
	g.mu.RLock()
	db := g.db
	driver := g.driver
	g.mu.RUnlock()

	if db == nil {
		g.setError(ErrG3DBConnectionNotOpen.String())
		return NewEmpty()
	}

	ctx := context.Background()
	var cancel context.CancelFunc
	if timeoutSeconds > 0 {
		ctx, cancel = context.WithTimeout(ctx, time.Duration(timeoutSeconds)*time.Second)
		defer cancel()
	}

	tx, err := db.BeginTx(ctx, &sql.TxOptions{ReadOnly: readOnly})
	if err != nil {
		g.setError(ErrG3DBTransactionFailed.String() + ": " + err.Error())
		return NewEmpty()
	}

	g.setError("")
	return g.vm.newG3DBTransactionValue(tx, driver)
}

func (g *G3DB) setMaxOpenConns(n int) {
	g.mu.RLock()
	db := g.db
	g.mu.RUnlock()
	if db != nil {
		db.SetMaxOpenConns(n)
	}
}

func (g *G3DB) setMaxIdleConns(n int) {
	g.mu.RLock()
	db := g.db
	g.mu.RUnlock()
	if db != nil {
		db.SetMaxIdleConns(n)
	}
}

func (g *G3DB) setConnMaxLifetime(seconds int) {
	g.mu.RLock()
	db := g.db
	g.mu.RUnlock()
	if db != nil {
		db.SetConnMaxLifetime(time.Duration(seconds) * time.Second)
	}
}

func (g *G3DB) setConnMaxIdleTime(seconds int) {
	g.mu.RLock()
	db := g.db
	g.mu.RUnlock()
	if db != nil {
		db.SetConnMaxIdleTime(time.Duration(seconds) * time.Second)
	}
}

// stats returns a Scripting.Dictionary containing sql.DBStats fields.
func (g *G3DB) stats() Value {
	g.mu.RLock()
	db := g.db
	g.mu.RUnlock()

	if db == nil {
		return NewEmpty()
	}

	s := db.Stats()
	dictVal := g.vm.newDictionaryObject()
	d := g.vm.dictionaryItems[dictVal.Num]
	d.itemSet(NewString("MaxOpenConnections"), NewInteger(int64(s.MaxOpenConnections)))
	d.itemSet(NewString("OpenConnections"), NewInteger(int64(s.OpenConnections)))
	d.itemSet(NewString("InUse"), NewInteger(int64(s.InUse)))
	d.itemSet(NewString("Idle"), NewInteger(int64(s.Idle)))
	d.itemSet(NewString("WaitCount"), NewInteger(s.WaitCount))
	d.itemSet(NewString("WaitDurationSeconds"), NewDouble(s.WaitDuration.Seconds()))
	d.itemSet(NewString("MaxIdleClosed"), NewInteger(s.MaxIdleClosed))
	d.itemSet(NewString("MaxIdleTimeClosed"), NewInteger(s.MaxIdleTimeClosed))
	d.itemSet(NewString("MaxLifetimeClosed"), NewInteger(s.MaxLifetimeClosed))
	return dictVal
}

// setError stores an error message under write-lock.
func (g *G3DB) setError(msg string) {
	g.mu.Lock()
	g.lastError = msg
	g.mu.Unlock()
}

// ---------------------------------------------------------------------------
// G3DBResultSet — Forward-Only Cursor
// ---------------------------------------------------------------------------

// G3DBResultSet wraps sql.Rows and exposes a forward-only cursor to VBScript.
// rowData holds the current row's column values as VM Value instances.
// fieldsID is the companion G3DBFields object registered in the VM.
type G3DBResultSet struct {
	vm       *VM
	driver   string
	rows     *sql.Rows
	columns  []string       // original column names (case-preserved)
	colIdx   map[string]int // lower-case column name → slice index
	rowData  []Value        // current row (len == len(columns))
	fieldsID int64          // VM ID of the companion G3DBFields object
	eof      bool
	bof      bool
	isClosed bool
	mu       sync.RWMutex
}

// newG3DBResultSetValue allocates a G3DBResultSet, its companion G3DBFields,
// advances to the first row, and returns the native-object handle.
func (vm *VM) newG3DBResultSetValue(rows *sql.Rows) Value {
	cols, _ := rows.Columns()
	colIdx := make(map[string]int, len(cols))
	for i, c := range cols {
		colIdx[strings.ToLower(c)] = i
	}

	rs := &G3DBResultSet{
		vm:      vm,
		rows:    rows,
		columns: cols,
		colIdx:  colIdx,
		rowData: make([]Value, len(cols)),
		eof:     false,
		bof:     true,
	}

	// Register the ResultSet.
	rsID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3dbResultSetItems[rsID] = rs

	// Register the companion Fields proxy.
	fieldsID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3dbFieldsItems[fieldsID] = &G3DBFields{rs: rs}
	rs.fieldsID = fieldsID

	// Advance to the first row immediately so BOF is false after creation.
	rs.moveNext()

	return Value{Type: VTNativeObject, Num: rsID}
}

// DispatchPropertyGet exposes EOF, BOF, Fields, and direct column-name access.
func (rs *G3DBResultSet) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(propertyName) {
	case "eof":
		rs.mu.RLock()
		v := rs.eof
		rs.mu.RUnlock()
		return NewBool(v)
	case "bof":
		rs.mu.RLock()
		v := rs.bof
		rs.mu.RUnlock()
		return NewBool(v)
	case "fields":
		return Value{Type: VTNativeObject, Num: rs.fieldsID}
	}

	// Fall through to direct field-name property access (e.g. rs.FieldName).
	rs.mu.RLock()
	idx, ok := rs.colIdx[strings.ToLower(propertyName)]
	var val Value
	if ok {
		val = rs.rowData[idx]
	}
	rs.mu.RUnlock()
	if ok {
		return val
	}
	return NewEmpty()
}

// DispatchPropertySet is a no-op for ResultSet (read-only cursor).
func (rs *G3DBResultSet) DispatchPropertySet(_ string, _ []Value) bool {
	return false
}

// DispatchMethod routes all VBScript method calls for G3DBResultSet.
// The empty method name ("") handles the default property: rs("fieldname").
func (rs *G3DBResultSet) DispatchMethod(methodName string, args []Value) Value {
	lower := strings.ToLower(methodName)

	// Default property call: rs("fieldname") or rs(0).
	if lower == "" {
		if len(args) == 0 {
			return NewEmpty()
		}
		key := args[0]
		rs.mu.RLock()
		defer rs.mu.RUnlock()
		// Numeric index.
		if key.Type == VTInteger || key.Type == VTDouble {
			idx := int(key.Num)
			if key.Type == VTDouble {
				idx = int(key.Flt)
			}
			if idx >= 0 && idx < len(rs.rowData) {
				return rs.rowData[idx]
			}
			return NewEmpty()
		}
		// Named access.
		if i, ok := rs.colIdx[strings.ToLower(key.String())]; ok {
			return rs.rowData[i]
		}
		return NewEmpty()
	}

	switch lower {
	// MoveNext() — advances the cursor to the next row. Sets EOF when exhausted.
	case "movenext":
		rs.moveNext()
		return NewEmpty()

	// Close() — releases the underlying sql.Rows.
	case "close":
		return NewBool(rs.closeRows())

	// GetRows([maxRows]) — returns all remaining rows as a 2D array [col][row].
	case "getrows":
		maxRows := -1
		if len(args) > 0 && (args[0].Type == VTInteger || args[0].Type == VTDouble) {
			maxRows = rs.vm.asInt(args[0])
		}
		return rs.getRows(maxRows)

	// Columns() — returns a 1D array of column name strings.
	case "columns":
		return rs.getColumnNames()
	}

	return NewEmpty()
}

// moveNext advances the cursor. It reads from sql.Rows under write-lock.
// After all rows are consumed, eof is set to true.
func (rs *G3DBResultSet) moveNext() {
	rs.mu.Lock()
	defer rs.mu.Unlock()

	if rs.isClosed || rs.rows == nil {
		rs.eof = true
		return
	}
	if !rs.rows.Next() {
		rs.eof = true
		return
	}

	rs.bof = false

	// sql.Scan requires []interface{} at the API boundary; convert immediately to []Value.
	dest := make([]interface{}, len(rs.columns))
	ptrs := make([]interface{}, len(rs.columns))
	for i := range dest {
		ptrs[i] = &dest[i]
	}
	if err := rs.rows.Scan(ptrs...); err != nil {
		rs.eof = true
		return
	}
	for i, raw := range dest {
		rs.rowData[i] = g3dbRawToValue(raw)
	}
}

// closeRows closes the underlying sql.Rows and marks the set as closed.
func (rs *G3DBResultSet) closeRows() bool {
	rs.mu.Lock()
	defer rs.mu.Unlock()

	if rs.isClosed || rs.rows == nil {
		return true
	}
	err := rs.rows.Close()
	rs.isClosed = true
	rs.rows = nil
	return err == nil
}

// getRows collects remaining rows into a 2D array [cols][rows] compatible with
// the Classic ASP Recordset.GetRows convention.
func (rs *G3DBResultSet) getRows(maxRows int) Value {
	rs.mu.Lock()
	defer rs.mu.Unlock()

	if rs.isClosed || rs.rows == nil {
		return NewEmpty()
	}

	cols := len(rs.columns)
	collected := make([][]Value, 0, 64)

	// Include current row if present.
	if !rs.eof && len(rs.rowData) == cols {
		row := make([]Value, cols)
		copy(row, rs.rowData)
		collected = append(collected, row)
	}

	// Read remaining rows.
	dest := make([]interface{}, cols)
	ptrs := make([]interface{}, cols)
	for i := range dest {
		ptrs[i] = &dest[i]
	}

	for rs.rows.Next() {
		if maxRows > 0 && len(collected) >= maxRows {
			break
		}
		if err := rs.rows.Scan(ptrs...); err != nil {
			break
		}
		row := make([]Value, cols)
		for i, raw := range dest {
			row[i] = g3dbRawToValue(raw)
		}
		collected = append(collected, row)
	}

	rs.eof = true
	rowCount := len(collected)
	if rowCount == 0 {
		return NewEmpty()
	}

	// Build [cols][rows] nested arrays.
	colArray := NewVBArray(0, cols)
	for c := 0; c < cols; c++ {
		rowArray := NewVBArray(0, rowCount)
		for r := 0; r < rowCount; r++ {
			rowArray.Values[r] = collected[r][c]
		}
		colArray.Values[c] = Value{Type: VTArray, Arr: rowArray}
	}
	return Value{Type: VTArray, Arr: colArray}
}

// getColumnNames returns a 1D Value array containing column name strings.
func (rs *G3DBResultSet) getColumnNames() Value {
	rs.mu.RLock()
	cols := rs.columns
	rs.mu.RUnlock()

	arr := make([]Value, len(cols))
	for i, c := range cols {
		arr[i] = NewString(c)
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, arr)}
}

// ---------------------------------------------------------------------------
// G3DBFields — Fields Collection Proxy
// ---------------------------------------------------------------------------

// G3DBFields proxies VBScript access to the current-row field values of a
// G3DBResultSet. It holds a live pointer to the parent ResultSet.
type G3DBFields struct {
	rs *G3DBResultSet
}

// DispatchPropertyGet resolves Count and direct item properties.
func (f *G3DBFields) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(propertyName) {
	case "count":
		f.rs.mu.RLock()
		n := len(f.rs.columns)
		f.rs.mu.RUnlock()
		return NewInteger(int64(n))
	}
	// Fallback: treat the property name as a field name.
	return f.DispatchMethod(propertyName, nil)
}

// DispatchPropertySet is a no-op (Fields collection is read-only).
func (f *G3DBFields) DispatchPropertySet(_ string, _ []Value) bool {
	return false
}

// DispatchMethod routes Fields method calls.
// The default property ("") and Item() both resolve a field by name or index.
func (f *G3DBFields) DispatchMethod(methodName string, args []Value) Value {
	lower := strings.ToLower(methodName)

	if lower == "" || lower == "item" {
		if len(args) == 0 {
			return NewEmpty()
		}
		key := args[0]
		f.rs.mu.RLock()
		defer f.rs.mu.RUnlock()
		// Numeric index.
		if key.Type == VTInteger || key.Type == VTDouble {
			idx := int(key.Num)
			if key.Type == VTDouble {
				idx = int(key.Flt)
			}
			if idx >= 0 && idx < len(f.rs.rowData) {
				return f.rs.rowData[idx]
			}
			return NewEmpty()
		}
		// Named access.
		if i, ok := f.rs.colIdx[strings.ToLower(key.String())]; ok {
			return f.rs.rowData[i]
		}
		return NewEmpty()
	}

	if lower == "count" {
		f.rs.mu.RLock()
		n := len(f.rs.columns)
		f.rs.mu.RUnlock()
		return NewInteger(int64(n))
	}

	return NewEmpty()
}

// ---------------------------------------------------------------------------
// G3DBRow — Single-Row Result (from QueryRow)
// ---------------------------------------------------------------------------

// G3DBRow wraps a sql.Row for single-row query results.
// Because sql.Row is one-shot, scanning consumes the row permanently.
type G3DBRow struct {
	vm  *VM
	row *sql.Row
}

// newG3DBRowValue registers a G3DBRow in the VM and returns its handle.
func (vm *VM) newG3DBRowValue(row *sql.Row) Value {
	obj := &G3DBRow{vm: vm, row: row}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3dbRowItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet is a no-op for G3DBRow; all access is through methods.
func (r *G3DBRow) DispatchPropertyGet(_ string) Value { return NewEmpty() }

// DispatchPropertySet is a no-op for G3DBRow.
func (r *G3DBRow) DispatchPropertySet(_ string, _ []Value) bool { return false }

// DispatchMethod routes VBScript method calls for G3DBRow.
func (r *G3DBRow) DispatchMethod(methodName string, args []Value) Value {
	switch strings.ToLower(methodName) {

	// Scan([colCount]) — scans the row into an array of colCount values.
	// If colCount is omitted (or 0), scans a single value and returns it directly.
	case "scan":
		if r.row == nil {
			return NewEmpty()
		}
		colCount := len(args)
		if colCount == 0 {
			// Single-value scan.
			var raw interface{}
			if err := r.row.Scan(&raw); err != nil {
				return NewEmpty()
			}
			r.row = nil
			return g3dbRawToValue(raw)
		}
		// Multi-value scan: build dest slice and convert.
		dest := make([]interface{}, colCount)
		ptrs := make([]interface{}, colCount)
		for i := range dest {
			ptrs[i] = &dest[i]
		}
		if err := r.row.Scan(ptrs...); err != nil {
			r.row = nil
			return NewEmpty()
		}
		r.row = nil
		vals := make([]Value, colCount)
		for i, raw := range dest {
			vals[i] = g3dbRawToValue(raw)
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, vals)}

	// ScanMap(col1, col2, ...) — scans the row into a named Scripting.Dictionary.
	case "scanmap":
		if r.row == nil || len(args) == 0 {
			return NewEmpty()
		}
		dest := make([]interface{}, len(args))
		ptrs := make([]interface{}, len(args))
		for i := range dest {
			ptrs[i] = &dest[i]
		}
		if err := r.row.Scan(ptrs...); err != nil {
			r.row = nil
			return NewEmpty()
		}
		r.row = nil
		dictVal := r.vm.newDictionaryObject()
		d := r.vm.dictionaryItems[dictVal.Num]
		for i, arg := range args {
			d.itemSet(NewString(arg.String()), g3dbRawToValue(dest[i]))
		}
		return dictVal
	}
	return NewEmpty()
}

// ---------------------------------------------------------------------------
// G3DBStatement — Prepared Statement
// ---------------------------------------------------------------------------

// G3DBStatement wraps a sql.Stmt for reuse across multiple executions.
type G3DBStatement struct {
	vm     *VM
	stmt   *sql.Stmt
	driver string
	closed bool
	mu     sync.RWMutex
}

// newG3DBStatementValue registers a G3DBStatement in the VM and returns its handle.
func (vm *VM) newG3DBStatementValue(stmt *sql.Stmt, driver string) Value {
	obj := &G3DBStatement{vm: vm, stmt: stmt, driver: driver}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3dbStatementItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet is a no-op for G3DBStatement.
func (s *G3DBStatement) DispatchPropertyGet(_ string) Value { return NewEmpty() }

// DispatchPropertySet is a no-op for G3DBStatement.
func (s *G3DBStatement) DispatchPropertySet(_ string, _ []Value) bool { return false }

// DispatchMethod routes VBScript method calls for the prepared statement.
func (s *G3DBStatement) DispatchMethod(methodName string, args []Value) Value {
	switch strings.ToLower(methodName) {

	// Query([params...]) — executes the prepared SELECT and returns a G3DBResultSet.
	case "query":
		s.mu.RLock()
		stmt := s.stmt
		closed := s.closed
		s.mu.RUnlock()
		if closed || stmt == nil {
			return NewEmpty()
		}
		iArgs := g3dbValueSliceToInterface(args)
		rows, err := stmt.Query(iArgs...)
		if err != nil {
			return NewEmpty()
		}
		return s.vm.newG3DBResultSetValue(rows)

	// QueryRow([params...]) — executes the prepared SELECT for one row.
	case "queryrow":
		s.mu.RLock()
		stmt := s.stmt
		closed := s.closed
		s.mu.RUnlock()
		if closed || stmt == nil {
			return NewEmpty()
		}
		iArgs := g3dbValueSliceToInterface(args)
		row := stmt.QueryRow(iArgs...)
		return s.vm.newG3DBRowValue(row)

	// Exec([params...]) — executes the prepared INSERT/UPDATE/DELETE.
	case "exec":
		s.mu.RLock()
		stmt := s.stmt
		closed := s.closed
		s.mu.RUnlock()
		if closed || stmt == nil {
			return NewEmpty()
		}
		iArgs := g3dbValueSliceToInterface(args)
		result, err := stmt.Exec(iArgs...)
		if err != nil {
			return NewEmpty()
		}
		return s.vm.newG3DBResultValue(result)

	// Close() — closes and releases the prepared statement.
	case "close":
		return NewBool(s.closeStmt())
	}
	return NewEmpty()
}

// closeStmt closes the underlying sql.Stmt in a thread-safe manner.
func (s *G3DBStatement) closeStmt() bool {
	s.mu.Lock()
	defer s.mu.Unlock()
	if s.closed || s.stmt == nil {
		return true
	}
	err := s.stmt.Close()
	s.closed = true
	s.stmt = nil
	return err == nil
}

// ---------------------------------------------------------------------------
// G3DBTransaction — Database Transaction
// ---------------------------------------------------------------------------

// G3DBTransaction wraps a sql.Tx, providing Commit, Rollback, and query methods.
// If the script ends without an explicit Commit, the transaction is automatically
// rolled back in ensureRollback to prevent resource leaks.
type G3DBTransaction struct {
	vm        *VM
	tx        *sql.Tx
	driver    string
	committed bool
	closed    bool
	mu        sync.RWMutex
}

// newG3DBTransactionValue registers a G3DBTransaction in the VM and returns its handle.
func (vm *VM) newG3DBTransactionValue(tx *sql.Tx, driver string) Value {
	obj := &G3DBTransaction{vm: vm, tx: tx, driver: driver}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3dbTransactionItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet exposes Committed and Closed transaction state.
func (t *G3DBTransaction) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(propertyName) {
	case "committed":
		t.mu.RLock()
		v := t.committed
		t.mu.RUnlock()
		return NewBool(v)
	case "closed":
		t.mu.RLock()
		v := t.closed
		t.mu.RUnlock()
		return NewBool(v)
	}
	return NewEmpty()
}

// DispatchPropertySet is a no-op for G3DBTransaction.
func (t *G3DBTransaction) DispatchPropertySet(_ string, _ []Value) bool { return false }

// DispatchMethod routes VBScript method calls for the transaction object.
func (t *G3DBTransaction) DispatchMethod(methodName string, args []Value) Value {
	switch strings.ToLower(methodName) {

	// Commit / CommitTrans — commits the transaction.
	case "commit", "committrans":
		return NewBool(t.commit())

	// Rollback / RollbackTrans — rolls back the transaction.
	case "rollback", "rollbacktrans":
		return NewBool(t.rollback())

	// Query(sql [, params...]) — executes a SELECT within the transaction.
	case "query":
		if len(args) < 1 {
			return NewEmpty()
		}
		return t.query(args[0].String(), args[1:])

	// QueryRow(sql [, params...]) — executes a single-row SELECT within the transaction.
	case "queryrow":
		if len(args) < 1 {
			return NewEmpty()
		}
		return t.queryRow(args[0].String(), args[1:])

	// Exec(sql [, params...]) — executes an INSERT/UPDATE/DELETE within the transaction.
	case "exec":
		if len(args) < 1 {
			return NewEmpty()
		}
		return t.exec(args[0].String(), args[1:])

	// Prepare(sql) — creates a prepared statement scoped to this transaction.
	case "prepare":
		if len(args) < 1 {
			return NewEmpty()
		}
		return t.prepare(args[0].String())
	}
	return NewEmpty()
}

func (t *G3DBTransaction) commit() bool {
	t.mu.Lock()
	defer t.mu.Unlock()
	if t.closed || t.tx == nil {
		return false
	}
	err := t.tx.Commit()
	t.committed = (err == nil)
	t.closed = true
	t.tx = nil
	return err == nil
}

func (t *G3DBTransaction) rollback() bool {
	t.mu.Lock()
	defer t.mu.Unlock()
	if t.closed || t.tx == nil {
		return false
	}
	err := t.tx.Rollback()
	t.closed = true
	t.tx = nil
	return err == nil
}

func (t *G3DBTransaction) query(sqlText string, params []Value) Value {
	t.mu.RLock()
	tx := t.tx
	driver := t.driver
	closed := t.closed
	t.mu.RUnlock()
	if closed || tx == nil {
		return NewEmpty()
	}
	prepared := g3dbRewritePlaceholders(sqlText, driver)
	rows, err := tx.Query(prepared, g3dbValueSliceToInterface(params)...)
	if err != nil {
		return NewEmpty()
	}
	return t.vm.newG3DBResultSetValue(rows)
}

func (t *G3DBTransaction) queryRow(sqlText string, params []Value) Value {
	t.mu.RLock()
	tx := t.tx
	driver := t.driver
	closed := t.closed
	t.mu.RUnlock()
	if closed || tx == nil {
		return NewEmpty()
	}
	prepared := g3dbRewritePlaceholders(sqlText, driver)
	row := tx.QueryRow(prepared, g3dbValueSliceToInterface(params)...)
	return t.vm.newG3DBRowValue(row)
}

func (t *G3DBTransaction) exec(sqlText string, params []Value) Value {
	t.mu.RLock()
	tx := t.tx
	driver := t.driver
	closed := t.closed
	t.mu.RUnlock()
	if closed || tx == nil {
		return NewEmpty()
	}
	prepared := g3dbRewritePlaceholders(sqlText, driver)
	result, err := tx.Exec(prepared, g3dbValueSliceToInterface(params)...)
	if err != nil {
		return NewEmpty()
	}
	return t.vm.newG3DBResultValue(result)
}

func (t *G3DBTransaction) prepare(sqlText string) Value {
	t.mu.RLock()
	tx := t.tx
	driver := t.driver
	closed := t.closed
	t.mu.RUnlock()
	if closed || tx == nil {
		return NewEmpty()
	}
	prepared := g3dbRewritePlaceholders(sqlText, driver)
	stmt, err := tx.Prepare(prepared)
	if err != nil {
		return NewEmpty()
	}
	return t.vm.newG3DBStatementValue(stmt, driver)
}

// ensureRollback automatically rolls back an uncommitted transaction. This is
// called implicitly when the VM map entry is iterated during cleanup.
func (t *G3DBTransaction) ensureRollback() {
	t.mu.Lock()
	defer t.mu.Unlock()
	if !t.closed && !t.committed && t.tx != nil {
		_ = t.tx.Rollback()
		t.closed = true
		t.tx = nil
	}
}

// ---------------------------------------------------------------------------
// G3DBResult — Exec Result Metadata
// ---------------------------------------------------------------------------

// G3DBResult wraps sql.Result to expose LastInsertId and RowsAffected to VBScript.
type G3DBResult struct {
	result sql.Result
}

// newG3DBResultValue registers a G3DBResult in the VM and returns its handle.
func (vm *VM) newG3DBResultValue(result sql.Result) Value {
	obj := &G3DBResult{result: result}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3dbResultItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet exposes LastInsertId and RowsAffected as properties.
func (r *G3DBResult) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(propertyName) {
	case "lastinsertid":
		id, _ := r.result.LastInsertId()
		return NewInteger(id)
	case "rowsaffected":
		rows, _ := r.result.RowsAffected()
		return NewInteger(rows)
	}
	return r.DispatchMethod(propertyName, nil)
}

// DispatchPropertySet is a no-op for G3DBResult.
func (r *G3DBResult) DispatchPropertySet(_ string, _ []Value) bool { return false }

// DispatchMethod exposes LastInsertId() and RowsAffected() as callable methods.
func (r *G3DBResult) DispatchMethod(methodName string, _ []Value) Value {
	switch strings.ToLower(methodName) {
	case "lastinsertid":
		id, _ := r.result.LastInsertId()
		return NewInteger(id)
	case "rowsaffected":
		rows, _ := r.result.RowsAffected()
		return NewInteger(rows)
	}
	return NewEmpty()
}

// ---------------------------------------------------------------------------
// Helper — SQL raw value → VM Value (the only interface{} boundary in G3DB)
// ---------------------------------------------------------------------------

// g3dbRawToValue converts a value returned by sql.Rows.Scan into a VM Value.
// interface{} is confined here at the SQL driver boundary and is not used elsewhere.
func g3dbRawToValue(raw interface{}) Value {
	if raw == nil {
		return NewNull()
	}
	switch v := raw.(type) {
	case bool:
		return NewBool(v)
	case int64:
		return NewInteger(v)
	case int32:
		return NewInteger(int64(v))
	case int:
		return NewInteger(int64(v))
	case float64:
		return NewDouble(v)
	case float32:
		return NewDouble(float64(v))
	case []byte:
		return NewString(string(v))
	case string:
		return NewString(v)
	case time.Time:
		return NewDate(v)
	default:
		return NewString(fmt.Sprintf("%v", v))
	}
}

// g3dbValueSliceToInterface converts a []Value parameter list into []interface{}
// for passing to sql.DB/Tx/Stmt methods. This is the only other required
// interface{} boundary — dictated by the database/sql API.
func g3dbValueSliceToInterface(vals []Value) []interface{} {
	if len(vals) == 0 {
		return nil
	}
	out := make([]interface{}, len(vals))
	for i, v := range vals {
		switch v.Type {
		case VTEmpty, VTNull:
			out[i] = nil
		case VTBool:
			out[i] = v.Num != 0
		case VTInteger:
			out[i] = v.Num
		case VTDouble:
			out[i] = v.Flt
		case VTDate:
			out[i] = time.Unix(0, v.Num).UTC()
		default:
			out[i] = v.String()
		}
	}
	return out
}

// ---------------------------------------------------------------------------
// Helper — Placeholder rewriting for multi-driver parameterized queries
// ---------------------------------------------------------------------------

// g3dbRewritePlaceholders rewrites ? placeholders in an SQL string to the
// per-driver positional placeholder syntax. MySQL and SQLite accept ? natively
// so no rewriting is needed for those drivers.
func g3dbRewritePlaceholders(sqlText, driver string) string {
	switch driver {
	case "postgres":
		// Replace ? with $1, $2, ... (PostgreSQL positional notation).
		return g3dbReplaceQuestionMarks(sqlText, func(n int) string {
			return fmt.Sprintf("$%d", n)
		})
	case "mssql":
		// Replace ? with @p1, @p2, ... (SQL Server named positional notation).
		return g3dbReplaceQuestionMarks(sqlText, func(n int) string {
			return fmt.Sprintf("@p%d", n)
		})
	case "oracle":
		// Replace ? with :p1, :p2, ... (go-ora/v2 positional notation).
		return g3dbReplaceQuestionMarks(sqlText, func(n int) string {
			return fmt.Sprintf(":p%d", n)
		})
	default:
		// mysql, sqlite — ? is the native placeholder; return unchanged.
		return sqlText
	}
}

// g3dbReplaceQuestionMarks iterates through the SQL string and replaces each
// unquoted ? with the string produced by replacer(n), where n increments from 1.
// Single-quoted, double-quoted, and back-tick quoted literals are skipped.
func g3dbReplaceQuestionMarks(sqlText string, replacer func(int) string) string {
	var sb strings.Builder
	sb.Grow(len(sqlText) + 16)
	n := 1
	in := []rune(sqlText)
	for i := 0; i < len(in); i++ {
		ch := in[i]
		switch ch {
		case '\'', '"', '`':
			// Copy quoted literal verbatim.
			quote := ch
			sb.WriteRune(ch)
			i++
			for i < len(in) {
				c := in[i]
				sb.WriteRune(c)
				if c == quote {
					break
				}
				i++
			}
		case '?':
			sb.WriteString(replacer(n))
			n++
		default:
			sb.WriteRune(ch)
		}
	}
	return sb.String()
}

// ---------------------------------------------------------------------------
// Helper — Driver name normalization
// ---------------------------------------------------------------------------

// g3dbNormalizeDriver converts common driver aliases to the canonical name
// expected by database/sql. Returns an empty string for unknown drivers.
func g3dbNormalizeDriver(driver string) string {
	switch strings.ToLower(strings.TrimSpace(driver)) {
	case "mysql", "mariadb":
		return "mysql"
	case "postgres", "postgresql", "pgsql":
		return "postgres"
	case "mssql", "sqlserver", "sql server":
		return "mssql"
	case "sqlite", "sqlite3":
		return "sqlite"
	case "oracle", "ora", "oci", "oci8":
		return "oracle"
	}
	return ""
}

// ---------------------------------------------------------------------------
// Helper — Viper config reader for OpenFromEnv
// ---------------------------------------------------------------------------

// g3dbNewConfigViper creates a Viper instance loaded with axonasp.toml and
// activates environment variable overrides when global.viper_automatic_env is true.
func g3dbNewConfigViper() *viper.Viper {
	return axonconfig.NewViper()
}

// g3dbBuildDSN constructs a driver-specific DSN string from the [g3db] section
// of axonasp.toml (or environment variable overrides). Returns an empty string
// when required keys are absent or the driver is not supported.
func g3dbBuildDSN(v *viper.Viper, driver string) string {
	switch driver {
	case "mysql":
		host := v.GetString("g3db.mysql_host")
		port := v.GetInt("g3db.mysql_port")
		user := v.GetString("g3db.mysql_user")
		pass := v.GetString("g3db.mysql_pass")
		dbName := v.GetString("g3db.mysql_database")
		if host == "" || user == "" || dbName == "" {
			return ""
		}
		if port == 0 {
			port = 3306
		}
		return fmt.Sprintf("%s:%s@tcp(%s:%d)/%s?parseTime=true", user, pass, host, port, dbName)

	case "postgres":
		host := v.GetString("g3db.postgres_host")
		port := v.GetInt("g3db.postgres_port")
		user := v.GetString("g3db.postgres_user")
		pass := v.GetString("g3db.postgres_pass")
		dbName := v.GetString("g3db.postgres_database")
		sslMode := v.GetString("g3db.postgres_ssl_mode")
		if host == "" || user == "" || dbName == "" {
			return ""
		}
		if port == 0 {
			port = 5432
		}
		if sslMode == "" {
			sslMode = "disable"
		}
		return fmt.Sprintf("host=%s port=%d user=%s password=%s dbname=%s sslmode=%s",
			host, port, user, pass, dbName, sslMode)

	case "mssql":
		host := v.GetString("g3db.mssql_host")
		port := v.GetInt("g3db.mssql_port")
		user := v.GetString("g3db.mssql_user")
		pass := v.GetString("g3db.mssql_pass")
		dbName := v.GetString("g3db.mssql_database")
		if host == "" || user == "" || dbName == "" {
			return ""
		}
		if port == 0 {
			port = 1433
		}
		return fmt.Sprintf("server=%s;port=%d;user id=%s;password=%s;database=%s",
			host, port, user, pass, dbName)

	case "sqlite":
		path := v.GetString("g3db.sqlite_path")
		if path == "" {
			return ""
		}
		timeout := v.GetInt("g3db.sqlite_busy_timeout")
		if timeout == 0 {
			timeout = 5000
		}
		return fmt.Sprintf("%s?_busy_timeout=%d", path, timeout)

	case "oracle":
		// go-ora/v2 accepts a URL-style DSN:
		//   oracle://user:pass@host:port/service_name
		// or a full connection string via g3db.oracle_dsn for advanced cases.
		dsn := v.GetString("g3db.oracle_dsn")
		if dsn != "" {
			return dsn
		}
		host := v.GetString("g3db.oracle_host")
		port := v.GetInt("g3db.oracle_port")
		user := v.GetString("g3db.oracle_user")
		pass := v.GetString("g3db.oracle_pass")
		service := v.GetString("g3db.oracle_service")
		if host == "" || user == "" || service == "" {
			return ""
		}
		if port == 0 {
			port = 1521
		}
		return fmt.Sprintf("oracle://%s:%s@%s:%d/%s", user, pass, host, port, service)
	}
	return ""
}
