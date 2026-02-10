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
package server

import (
	"context"
	"database/sql"
	"fmt"
	"os"
	"strings"
	"sync"
	"time"

	_ "github.com/denisenkom/go-mssqldb"
	_ "github.com/go-sql-driver/mysql"
	_ "github.com/lib/pq"
	_ "modernc.org/sqlite"
)

// --- G3DB Main Object ---

// G3DB provides direct access to Go's database/sql functionality with VBScript compatibility
type G3DB struct {
	db          *sql.DB
	ctx         *ExecutionContext
	driver      string
	dsn         string
	isOpen      bool
	mu          sync.RWMutex
	lastError   error
}

// NewG3DB creates a new G3DB instance
func NewG3DB(ctx *ExecutionContext) *G3DB {
	g3db := &G3DB{
		ctx:    ctx,
		isOpen: false,
	}
	// Register for automatic cleanup
	if ctx != nil {
		ctx.RegisterManagedResource(g3db)
	}
	return g3db
}

func (g *G3DB) GetProperty(name string) interface{} {
	g.mu.RLock()
	defer g.mu.RUnlock()

	switch strings.ToLower(name) {
	case "isopen":
		return g.isOpen
	case "driver":
		return g.driver
	case "dsn":
		return g.dsn
	case "lasterror":
		if g.lastError != nil {
			return g.lastError.Error()
		}
		return ""
	}
	return nil
}

func (g *G3DB) SetProperty(name string, value interface{}) {
	g.mu.Lock()
	defer g.mu.Unlock()

	switch strings.ToLower(name) {
	case "driver":
		g.driver = fmt.Sprintf("%v", value)
	case "dsn":
		g.dsn = fmt.Sprintf("%v", value)
	}
}

func (g *G3DB) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	switch method {
	case "open":
		// Open(driver, dsn)
		if len(args) < 2 {
			g.setError(fmt.Errorf("open requires driver and dsn parameters"))
			return false
		}
		driver := fmt.Sprintf("%v", args[0])
		dsn := fmt.Sprintf("%v", args[1])
		return g.open(driver, dsn)

	case "openfromenv":
		// OpenFromEnv([driver]) - uses .env configuration
		driver := "mysql" // default
		if len(args) > 0 {
			driver = fmt.Sprintf("%v", args[0])
		}
		return g.openFromEnv(driver)

	case "close":
		return g.close()

	case "query":
		// Query(sql, [params...])
		if len(args) < 1 {
			g.setError(fmt.Errorf("query requires sql parameter"))
			return nil
		}
		sqlText := fmt.Sprintf("%v", args[0])
		params := normalizeExecuteParams(args[1:])
		return g.query(sqlText, params...)

	case "queryrow":
		// QueryRow(sql, [params...])
		if len(args) < 1 {
			g.setError(fmt.Errorf("queryrow requires sql parameter"))
			return nil
		}
		sqlText := fmt.Sprintf("%v", args[0])
		params := normalizeExecuteParams(args[1:])
		return g.queryRow(sqlText, params...)

	case "exec":
		// Exec(sql, [params...])
		if len(args) < 1 {
			g.setError(fmt.Errorf("exec requires sql parameter"))
			return nil
		}
		sqlText := fmt.Sprintf("%v", args[0])
		params := normalizeExecuteParams(args[1:])
		return g.exec(sqlText, params...)

	case "prepare":
		// Prepare(sql)
		if len(args) < 1 {
			g.setError(fmt.Errorf("prepare requires sql parameter"))
			return nil
		}
		sqlText := fmt.Sprintf("%v", args[0])
		return g.prepare(sqlText)

	case "preparecontext":
		// PrepareContext(timeout, sql) - timeout in seconds
		if len(args) < 2 {
			g.setError(fmt.Errorf("preparecontext requires timeout and sql parameters"))
			return nil
		}
		timeout := toInt(args[0])
		sqlText := fmt.Sprintf("%v", args[1])
		return g.prepareContext(timeout, sqlText)

	case "begin", "begintrans", "begintransaction":
		// Begin()
		return g.beginTx()

	case "begintx":
		// BeginTx(timeout, readOnly) - timeout in seconds
		timeout := 0
		readOnly := false
		if len(args) > 0 {
			timeout = toInt(args[0])
		}
		if len(args) > 1 {
			readOnly = toBool(args[1])
		}
		return g.beginTxWithOptions(timeout, readOnly)

	case "setmaxopenconns":
		// SetMaxOpenConns(n)
		if len(args) < 1 {
			return false
		}
		n := toInt(args[0])
		g.setMaxOpenConns(n)
		return true

	case "setmaxidleconns":
		// SetMaxIdleConns(n)
		if len(args) < 1 {
			return false
		}
		n := toInt(args[0])
		g.setMaxIdleConns(n)
		return true

	case "setconnmaxlifetime":
		// SetConnMaxLifetime(seconds)
		if len(args) < 1 {
			return false
		}
		seconds := toInt(args[0])
		g.setConnMaxLifetime(seconds)
		return true

	case "setconnmaxidletime":
		// SetConnMaxIdleTime(seconds)
		if len(args) < 1 {
			return false
		}
		seconds := toInt(args[0])
		g.setConnMaxIdleTime(seconds)
		return true

	case "stats":
		// Stats() - returns dictionary with connection stats
		return g.stats()

	case "geterror", "getlasterror":
		if g.lastError != nil {
			return g.lastError.Error()
		}
		return ""
	}

	return nil
}

// Cleanup ensures database is closed when context ends
func (g *G3DB) Cleanup() {
	g.close()
}

// open establishes database connection
func (g *G3DB) open(driver, dsn string) bool {
	g.mu.Lock()
	defer g.mu.Unlock()

	if g.isOpen {
		g.lastError = fmt.Errorf("connection already open")
		return false
	}

	// Normalize driver names
	driver = normalizeDriverName(driver)

	db, err := sql.Open(driver, dsn)
	if err != nil {
		g.lastError = err
		return false
	}

	// Test connection
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	if err := db.PingContext(ctx); err != nil {
		db.Close()
		g.lastError = err
		return false
	}

	g.db = db
	g.driver = driver
	g.dsn = dsn
	g.isOpen = true
	g.lastError = nil
	return true
}

// openFromEnv opens connection using .env configuration
func (g *G3DB) openFromEnv(driver string) bool {
	driver = normalizeDriverName(driver)

	var dsn string
	envPrefix := strings.ToUpper(driver)

	switch driver {
	case "mysql":
		host := getEnv(envPrefix+"_HOST", "localhost")
		port := getEnv(envPrefix+"_PORT", "3306")
		user := getEnv(envPrefix+"_USER", "root")
		pass := getEnv(envPrefix+"_PASS", "")
		database := getEnv(envPrefix+"_DATABASE", "test")
		dsn = fmt.Sprintf("%s:%s@tcp(%s:%s)/%s?parseTime=true", user, pass, host, port, database)

	case "postgres":
		host := getEnv(envPrefix+"_HOST", "localhost")
		port := getEnv(envPrefix+"_PORT", "5432")
		user := getEnv(envPrefix+"_USER", "postgres")
		pass := getEnv(envPrefix+"_PASS", "")
		database := getEnv(envPrefix+"_DATABASE", "test")
		sslmode := getEnv(envPrefix+"_SSLMODE", "disable")
		dsn = fmt.Sprintf("host=%s port=%s user=%s password=%s dbname=%s sslmode=%s", 
			host, port, user, pass, database, sslmode)

	case "mssql", "sqlserver":
		host := getEnv(envPrefix+"_HOST", "localhost")
		port := getEnv(envPrefix+"_PORT", "1433")
		user := getEnv(envPrefix+"_USER", "sa")
		pass := getEnv(envPrefix+"_PASS", "")
		database := getEnv(envPrefix+"_DATABASE", "test")
		dsn = fmt.Sprintf("server=%s;port=%s;user id=%s;password=%s;database=%s", 
			host, port, user, pass, database)

	case "sqlite":
		dbPath := getEnv(envPrefix+"_PATH", "./database.db")
		dsn = dbPath

	default:
		g.setError(fmt.Errorf("unsupported driver: %s", driver))
		return false
	}

	return g.open(driver, dsn)
}

// close closes database connection
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
		g.lastError = err
		return false
	}
	g.lastError = nil
	return true
}

// query executes a query and returns a result set
func (g *G3DB) query(sqlText string, params ...interface{}) interface{} {
	g.mu.RLock()
	db := g.db
	driver := g.driver
	g.mu.RUnlock()

	if db == nil {
		g.setError(fmt.Errorf("connection not open"))
		return nil
	}

	prepared := rewritePlaceholders(sqlText, driver)
	rows, err := db.Query(prepared, params...)
	if err != nil {
		g.setError(err)
		return nil
	}

	g.lastError = nil
	return NewG3DBResultSet(rows, g.ctx)
}

// queryRow executes a query that returns a single row
func (g *G3DB) queryRow(sqlText string, params ...interface{}) interface{} {
	g.mu.RLock()
	db := g.db
	driver := g.driver
	g.mu.RUnlock()

	if db == nil {
		g.setError(fmt.Errorf("connection not open"))
		return nil
	}

	prepared := rewritePlaceholders(sqlText, driver)
	row := db.QueryRow(prepared, params...)
	
	g.lastError = nil
	return NewG3DBRow(row, g.ctx)
}

// exec executes a command (INSERT, UPDATE, DELETE) and returns result info
func (g *G3DB) exec(sqlText string, params ...interface{}) interface{} {
	g.mu.RLock()
	db := g.db
	driver := g.driver
	g.mu.RUnlock()

	if db == nil {
		g.setError(fmt.Errorf("connection not open"))
		return nil
	}

	prepared := rewritePlaceholders(sqlText, driver)
	result, err := db.Exec(prepared, params...)
	if err != nil {
		g.setError(err)
		return nil
	}

	g.lastError = nil
	return NewG3DBResult(result)
}

// prepare prepares a statement
func (g *G3DB) prepare(sqlText string) interface{} {
	g.mu.RLock()
	db := g.db
	driver := g.driver
	g.mu.RUnlock()

	if db == nil {
		g.setError(fmt.Errorf("connection not open"))
		return nil
	}

	prepared := rewritePlaceholders(sqlText, driver)
	stmt, err := db.Prepare(prepared)
	if err != nil {
		g.setError(err)
		return nil
	}

	g.lastError = nil
	return NewG3DBStatement(stmt, g.ctx)
}

// prepareContext prepares a statement with context timeout
func (g *G3DB) prepareContext(timeoutSeconds int, sqlText string) interface{} {
	g.mu.RLock()
	db := g.db
	driver := g.driver
	g.mu.RUnlock()

	if db == nil {
		g.setError(fmt.Errorf("connection not open"))
		return nil
	}

	ctx, cancel := context.WithTimeout(context.Background(), time.Duration(timeoutSeconds)*time.Second)
	defer cancel()

	prepared := rewritePlaceholders(sqlText, driver)
	stmt, err := db.PrepareContext(ctx, prepared)
	if err != nil {
		g.setError(err)
		return nil
	}

	g.lastError = nil
	return NewG3DBStatement(stmt, g.ctx)
}

// beginTx starts a transaction
func (g *G3DB) beginTx() interface{} {
	g.mu.RLock()
	db := g.db
	g.mu.RUnlock()

	if db == nil {
		g.setError(fmt.Errorf("connection not open"))
		return nil
	}

	tx, err := db.Begin()
	if err != nil {
		g.setError(err)
		return nil
	}

	g.lastError = nil
	return NewG3DBTransaction(tx, g.driver, g.ctx)
}

// beginTxWithOptions starts a transaction with options
func (g *G3DB) beginTxWithOptions(timeoutSeconds int, readOnly bool) interface{} {
	g.mu.RLock()
	db := g.db
	driver := g.driver
	g.mu.RUnlock()

	if db == nil {
		g.setError(fmt.Errorf("connection not open"))
		return nil
	}

	ctx := context.Background()
	if timeoutSeconds > 0 {
		var cancel context.CancelFunc
		ctx, cancel = context.WithTimeout(ctx, time.Duration(timeoutSeconds)*time.Second)
		defer cancel()
	}

	opts := &sql.TxOptions{
		ReadOnly: readOnly,
	}

	tx, err := db.BeginTx(ctx, opts)
	if err != nil {
		g.setError(err)
		return nil
	}

	g.lastError = nil
	return NewG3DBTransaction(tx, driver, g.ctx)
}

// setMaxOpenConns sets maximum open connections
func (g *G3DB) setMaxOpenConns(n int) {
	g.mu.RLock()
	defer g.mu.RUnlock()
	if g.db != nil {
		g.db.SetMaxOpenConns(n)
	}
}

// setMaxIdleConns sets maximum idle connections
func (g *G3DB) setMaxIdleConns(n int) {
	g.mu.RLock()
	defer g.mu.RUnlock()
	if g.db != nil {
		g.db.SetMaxIdleConns(n)
	}
}

// setConnMaxLifetime sets maximum connection lifetime
func (g *G3DB) setConnMaxLifetime(seconds int) {
	g.mu.RLock()
	defer g.mu.RUnlock()
	if g.db != nil {
		g.db.SetConnMaxLifetime(time.Duration(seconds) * time.Second)
	}
}

// setConnMaxIdleTime sets maximum connection idle time
func (g *G3DB) setConnMaxIdleTime(seconds int) {
	g.mu.RLock()
	defer g.mu.RUnlock()
	if g.db != nil {
		g.db.SetConnMaxIdleTime(time.Duration(seconds) * time.Second)
	}
}

// stats returns connection statistics
func (g *G3DB) stats() interface{} {
	g.mu.RLock()
	defer g.mu.RUnlock()

	if g.db == nil {
		return nil
	}

	stats := g.db.Stats()
	
	// Return as VBScript-compatible dictionary
	dl := NewDictionary(g.ctx)
	dict := dl.dict
	dict.Add([]interface{}{"MaxOpenConnections", stats.MaxOpenConnections})
	dict.Add([]interface{}{"OpenConnections", stats.OpenConnections})
	dict.Add([]interface{}{"InUse", stats.InUse})
	dict.Add([]interface{}{"Idle", stats.Idle})
	dict.Add([]interface{}{"WaitCount", stats.WaitCount})
	dict.Add([]interface{}{"WaitDuration", stats.WaitDuration.Seconds()})
	dict.Add([]interface{}{"MaxIdleClosed", stats.MaxIdleClosed})
	dict.Add([]interface{}{"MaxIdleTimeClosed", stats.MaxIdleTimeClosed})
	dict.Add([]interface{}{"MaxLifetimeClosed", stats.MaxLifetimeClosed})

	return dl
}

func (g *G3DB) setError(err error) {
	g.mu.Lock()
	g.lastError = err
	g.mu.Unlock()
}

// --- G3DBResultSet ---

// G3DBResultSet wraps sql.Rows for VBScript compatibility
type G3DBResultSet struct {
	rows        *sql.Rows
	ctx         *ExecutionContext
	columns     []string
	currentData map[string]interface{}
	EOF         bool
	BOF         bool
	Fields      *FieldsCollection
	isClosed    bool
	mu          sync.RWMutex
}

// NewG3DBResultSet creates a new result set
func NewG3DBResultSet(rows *sql.Rows, ctx *ExecutionContext) *G3DBResultSet {
	columns, _ := rows.Columns()
	
	rs := &G3DBResultSet{
		rows:        rows,
		ctx:         ctx,
		columns:     columns,
		currentData: make(map[string]interface{}),
		EOF:         false,
		BOF:         true,
		Fields:      NewFieldsCollection(),
	}

	// Initialize fields collection
	for _, col := range columns {
		field := &Field{Name: col, Value: nil}
		rs.Fields.fields = append(rs.Fields.fields, field)
		rs.Fields.data[strings.ToLower(col)] = nil
	}

	// Move to first record
	rs.moveNext()

	// Register for cleanup
	if ctx != nil {
		ctx.RegisterManagedResource(rs)
	}

	return rs
}

func (rs *G3DBResultSet) GetProperty(name string) interface{} {
	rs.mu.RLock()
	defer rs.mu.RUnlock()

	switch strings.ToLower(name) {
	case "eof":
		return rs.EOF
	case "bof":
		return rs.BOF
	case "fields":
		return rs.Fields
	default:
		// Try to access as field name
		if val, ok := rs.currentData[strings.ToLower(name)]; ok {
			return val
		}
	}
	return nil
}

func (rs *G3DBResultSet) SetProperty(name string, value interface{}) {
	// ResultSet is read-only
}

func (rs *G3DBResultSet) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	// Default method returns field value
	if method == "" && len(args) > 0 {
		fieldName := strings.ToLower(fmt.Sprintf("%v", args[0]))
		rs.mu.RLock()
		val := rs.currentData[fieldName]
		rs.mu.RUnlock()
		return val
	}

	switch method {
	case "movenext":
		return rs.moveNext()
	case "close":
		return rs.closers()
	case "getrows":
		return rs.getRows()
	case "columns":
		return rs.getColumns()
	}

	return nil
}

func (rs *G3DBResultSet) moveNext() bool {
	rs.mu.Lock()
	defer rs.mu.Unlock()

	if rs.isClosed || rs.rows == nil {
		rs.EOF = true
		return false
	}

	if !rs.rows.Next() {
		rs.EOF = true
		return false
	}

	rs.BOF = false

	// Scan current row
	values := make([]interface{}, len(rs.columns))
	valuePtrs := make([]interface{}, len(rs.columns))
	for i := range values {
		valuePtrs[i] = &values[i]
	}

	if err := rs.rows.Scan(valuePtrs...); err != nil {
		rs.EOF = true
		return false
	}

	// Update current data and fields
	rs.currentData = make(map[string]interface{})
	for i, col := range rs.columns {
		val := convertSQLValue(values[i])
		rs.currentData[strings.ToLower(col)] = val
		rs.Fields.fields[i].Value = val
		rs.Fields.data[strings.ToLower(col)] = val
	}

	return true
}

func (rs *G3DBResultSet) closers() bool {
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

func (rs *G3DBResultSet) getRows() interface{} {
	rs.mu.Lock()
	defer rs.mu.Unlock()

	if rs.isClosed || rs.rows == nil {
		return NewVBArrayFromValues(0, []interface{}{})
	}

	allRows := make([]interface{}, 0)

	// Include current row if not at EOF
	if !rs.EOF && len(rs.currentData) > 0 {
		rowData := make(map[string]interface{})
		for k, v := range rs.currentData {
			rowData[k] = v
		}
		allRows = append(allRows, rowData)
	}

	// Read remaining rows
	for rs.rows.Next() {
		values := make([]interface{}, len(rs.columns))
		valuePtrs := make([]interface{}, len(rs.columns))
		for i := range values {
			valuePtrs[i] = &values[i]
		}

		if err := rs.rows.Scan(valuePtrs...); err != nil {
			break
		}

		rowData := make(map[string]interface{})
		for i, col := range rs.columns {
			rowData[strings.ToLower(col)] = convertSQLValue(values[i])
		}
		allRows = append(allRows, rowData)
	}

	rs.EOF = true
	return NewVBArrayFromValues(0, allRows)
}

func (rs *G3DBResultSet) getColumns() interface{} {
	rs.mu.RLock()
	defer rs.mu.RUnlock()

	cols := make([]interface{}, len(rs.columns))
	for i, col := range rs.columns {
		cols[i] = col
	}
	return NewVBArrayFromValues(0, cols)
}

func (rs *G3DBResultSet) Cleanup() {
	rs.closers()
}

// --- G3DBRow ---

// G3DBRow wraps sql.Row for single row queries
type G3DBRow struct {
	row *sql.Row
	ctx *ExecutionContext
}

func NewG3DBRow(row *sql.Row, ctx *ExecutionContext) *G3DBRow {
	return &G3DBRow{
		row: row,
		ctx: ctx,
	}
}

func (r *G3DBRow) GetProperty(name string) interface{} {
	return nil
}

func (r *G3DBRow) SetProperty(name string, value interface{}) {
}

func (r *G3DBRow) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	switch method {
	case "scan":
		// Scan(dest1, dest2, ...) - returns array of scanned values
		if r.row == nil {
			return nil
		}

		numCols := len(args)
		if numCols == 0 {
			// No args provided, try to scan into a single value
			var val interface{}
			if err := r.row.Scan(&val); err != nil {
				return nil
			}
			return convertSQLValue(val)
		}

		// Scan into provided number of columns
		values := make([]interface{}, numCols)
		valuePtrs := make([]interface{}, numCols)
		for i := range values {
			valuePtrs[i] = &values[i]
		}

		if err := r.row.Scan(valuePtrs...); err != nil {
			return nil
		}

		// Convert and return as array
		result := make([]interface{}, numCols)
		for i, val := range values {
			result[i] = convertSQLValue(val)
		}
		return NewVBArrayFromValues(0, result)

	case "scanmap":
		// ScanMap(column1, column2, ...) - returns dictionary with named values
		if r.row == nil {
			return nil
		}

		columnNames := make([]string, len(args))
		for i, arg := range args {
			columnNames[i] = fmt.Sprintf("%v", arg)
		}

		values := make([]interface{}, len(columnNames))
		valuePtrs := make([]interface{}, len(columnNames))
		for i := range values {
			valuePtrs[i] = &values[i]
		}

		if err := r.row.Scan(valuePtrs...); err != nil {
			return nil
		}

		dict := NewDictionary(r.ctx)
		d := dict.dict
		for i, col := range columnNames {
			d.Add([]interface{}{col, convertSQLValue(values[i])})
		}
		return dict
	}

	return nil
}

// --- G3DBStatement ---

// G3DBStatement wraps sql.Stmt for prepared statements
type G3DBStatement struct {
	stmt   *sql.Stmt
	ctx    *ExecutionContext
	mu     sync.RWMutex
	closed bool
}

func NewG3DBStatement(stmt *sql.Stmt, ctx *ExecutionContext) *G3DBStatement {
	s := &G3DBStatement{
		stmt:   stmt,
		ctx:    ctx,
		closed: false,
	}
	if ctx != nil {
		ctx.RegisterManagedResource(s)
	}
	return s
}

func (s *G3DBStatement) GetProperty(name string) interface{} {
	return nil
}

func (s *G3DBStatement) SetProperty(name string, value interface{}) {
}

func (s *G3DBStatement) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	switch method {
	case "query":
		// Query([params...])
		params := normalizeExecuteParams(args)
		return s.query(params...)

	case "queryrow":
		// QueryRow([params...])
		params := normalizeExecuteParams(args)
		return s.queryRow(params...)

	case "exec":
		// Exec([params...])
		params := normalizeExecuteParams(args)
		return s.exec(params...)

	case "close":
		return s.closeStmt()
	}

	return nil
}

func (s *G3DBStatement) query(params ...interface{}) interface{} {
	s.mu.RLock()
	stmt := s.stmt
	s.mu.RUnlock()

	if stmt == nil || s.closed {
		return nil
	}

	rows, err := stmt.Query(params...)
	if err != nil {
		return nil
	}

	return NewG3DBResultSet(rows, s.ctx)
}

func (s *G3DBStatement) queryRow(params ...interface{}) interface{} {
	s.mu.RLock()
	stmt := s.stmt
	s.mu.RUnlock()

	if stmt == nil || s.closed {
		return nil
	}

	row := stmt.QueryRow(params...)
	return NewG3DBRow(row, s.ctx)
}

func (s *G3DBStatement) exec(params ...interface{}) interface{} {
	s.mu.RLock()
	stmt := s.stmt
	s.mu.RUnlock()

	if stmt == nil || s.closed {
		return nil
	}

	result, err := stmt.Exec(params...)
	if err != nil {
		return nil
	}

	return NewG3DBResult(result)
}

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

func (s *G3DBStatement) Cleanup() {
	s.closeStmt()
}

// --- G3DBTransaction ---

// G3DBTransaction wraps sql.Tx for transactions
type G3DBTransaction struct {
	tx        *sql.Tx
	driver    string
	ctx       *ExecutionContext
	mu        sync.RWMutex
	committed bool
	closed    bool
}

func NewG3DBTransaction(tx *sql.Tx, driver string, ctx *ExecutionContext) *G3DBTransaction {
	t := &G3DBTransaction{
		tx:        tx,
		driver:    driver,
		ctx:       ctx,
		committed: false,
		closed:    false,
	}
	if ctx != nil {
		ctx.RegisterManagedResource(t)
	}
	return t
}

func (t *G3DBTransaction) GetProperty(name string) interface{} {
	t.mu.RLock()
	defer t.mu.RUnlock()

	switch strings.ToLower(name) {
	case "committed":
		return t.committed
	case "closed":
		return t.closed
	}
	return nil
}

func (t *G3DBTransaction) SetProperty(name string, value interface{}) {
}

func (t *G3DBTransaction) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	switch method {
	case "commit", "committrans":
		return t.commit()

	case "rollback", "rollbacktrans":
		return t.rollback()

	case "query":
		// Query(sql, [params...])
		if len(args) < 1 {
			return nil
		}
		sqlText := fmt.Sprintf("%v", args[0])
		params := normalizeExecuteParams(args[1:])
		return t.query(sqlText, params...)

	case "queryrow":
		// QueryRow(sql, [params...])
		if len(args) < 1 {
			return nil
		}
		sqlText := fmt.Sprintf("%v", args[0])
		params := normalizeExecuteParams(args[1:])
		return t.queryRow(sqlText, params...)

	case "exec":
		// Exec(sql, [params...])
		if len(args) < 1 {
			return nil
		}
		sqlText := fmt.Sprintf("%v", args[0])
		params := normalizeExecuteParams(args[1:])
		return t.exec(sqlText, params...)

	case "prepare":
		// Prepare(sql)
		if len(args) < 1 {
			return nil
		}
		sqlText := fmt.Sprintf("%v", args[0])
		return t.prepare(sqlText)
	}

	return nil
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
	return err == nil
}

func (t *G3DBTransaction) query(sqlText string, params ...interface{}) interface{} {
	t.mu.RLock()
	tx := t.tx
	driver := t.driver
	t.mu.RUnlock()

	if tx == nil || t.closed {
		return nil
	}

	prepared := rewritePlaceholders(sqlText, driver)
	rows, err := tx.Query(prepared, params...)
	if err != nil {
		return nil
	}

	return NewG3DBResultSet(rows, t.ctx)
}

func (t *G3DBTransaction) queryRow(sqlText string, params ...interface{}) interface{} {
	t.mu.RLock()
	tx := t.tx
	driver := t.driver
	t.mu.RUnlock()

	if tx == nil || t.closed {
		return nil
	}

	prepared := rewritePlaceholders(sqlText, driver)
	row := tx.QueryRow(prepared, params...)
	return NewG3DBRow(row, t.ctx)
}

func (t *G3DBTransaction) exec(sqlText string, params ...interface{}) interface{} {
	t.mu.RLock()
	tx := t.tx
	driver := t.driver
	t.mu.RUnlock()

	if tx == nil || t.closed {
		return nil
	}

	prepared := rewritePlaceholders(sqlText, driver)
	result, err := tx.Exec(prepared, params...)
	if err != nil {
		return nil
	}

	return NewG3DBResult(result)
}

func (t *G3DBTransaction) prepare(sqlText string) interface{} {
	t.mu.RLock()
	tx := t.tx
	driver := t.driver
	t.mu.RUnlock()

	if tx == nil || t.closed {
		return nil
	}

	prepared := rewritePlaceholders(sqlText, driver)
	stmt, err := tx.Prepare(prepared)
	if err != nil {
		return nil
	}

	return NewG3DBStatement(stmt, t.ctx)
}

func (t *G3DBTransaction) Cleanup() {
	t.mu.Lock()
	defer t.mu.Unlock()

	if !t.closed && t.tx != nil && !t.committed {
		// Auto-rollback if not committed
		t.tx.Rollback()
		t.closed = true
	}
}

// --- G3DBResult ---

// G3DBResult wraps sql.Result for exec results
type G3DBResult struct {
	result sql.Result
}

func NewG3DBResult(result sql.Result) *G3DBResult {
	return &G3DBResult{result: result}
}

func (r *G3DBResult) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "lastinsertid":
		id, _ := r.result.LastInsertId()
		return id
	case "rowsaffected":
		rows, _ := r.result.RowsAffected()
		return rows
	}
	return nil
}

func (r *G3DBResult) SetProperty(name string, value interface{}) {
}

func (r *G3DBResult) CallMethod(name string, args ...interface{}) interface{} {
	switch strings.ToLower(name) {
	case "lastinsertid":
		id, _ := r.result.LastInsertId()
		return id
	case "rowsaffected":
		rows, _ := r.result.RowsAffected()
		return rows
	}
	return nil
}

// --- Helper Functions ---

// normalizeDriverName normalizes database driver names
func normalizeDriverName(driver string) string {
	driver = strings.ToLower(strings.TrimSpace(driver))
	switch driver {
	case "mysql", "mariadb":
		return "mysql"
	case "postgres", "postgresql", "pgsql":
		return "postgres"
	case "mssql", "sqlserver", "sql server":
		return "mssql"
	case "sqlite", "sqlite3":
		return "sqlite"
	}
	return driver
}

// convertSQLValue converts SQL values to VBScript-compatible types
func convertSQLValue(val interface{}) interface{} {
	if val == nil {
		return nil
	}

	switch v := val.(type) {
	case []byte:
		return string(v)
	case time.Time:
		return v
	case int64:
		return int(v)
	default:
		return v
	}
}

// getEnv gets environment variable with default value
func getEnv(key, defaultValue string) string {
	if value := os.Getenv(key); value != "" {
		return value
	}
	return defaultValue
}

// toBool converts value to boolean
func toBool(val interface{}) bool {
	if val == nil {
		return false
	}
	switch v := val.(type) {
	case bool:
		return v
	case int, int32, int64:
		return v != 0
	case float64:
		return v != 0
	case string:
		s := strings.ToLower(v)
		return s == "true" || s == "yes" || s == "1"
	}
	return false
}
