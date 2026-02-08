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
	"database/sql"
	"fmt"
	"runtime"
	"sort"
	"strings"
	"time"

	_ "github.com/denisenkom/go-mssqldb"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	_ "github.com/go-sql-driver/mysql"
	_ "github.com/lib/pq"
	_ "modernc.org/sqlite"
)

// --- ADODB.Error ---

// ADOError represents a single database error
type ADOError struct {
	Number      int
	Description string
	Source      string
	SQLState    string
}

func (e *ADOError) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "number":
		return e.Number
	case "description":
		return e.Description
	case "source":
		return e.Source
	case "sqlstate":
		return e.SQLState
	}
	return nil
}

func (e *ADOError) SetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "number":
		e.Number = toInt(value)
	case "description":
		e.Description = fmt.Sprintf("%v", value)
	case "source":
		e.Source = fmt.Sprintf("%v", value)
	case "sqlstate":
		e.SQLState = fmt.Sprintf("%v", value)
	}
	return nil
}

func (e *ADOError) CallMethod(name string, args ...interface{}) interface{} {
	return nil
}

// ErrorsCollection holds a collection of ADOError objects
type ErrorsCollection struct {
	errors []*ADOError
}

func NewErrorsCollection() *ErrorsCollection {
	return &ErrorsCollection{
		errors: make([]*ADOError, 0),
	}
}

func (ec *ErrorsCollection) AddError(number int, description, source, sqlState string) {
	ec.errors = append(ec.errors, &ADOError{
		Number:      number,
		Description: description,
		Source:      source,
		SQLState:    sqlState,
	})
}

func (ec *ErrorsCollection) Clear() {
	ec.errors = make([]*ADOError, 0)
}

func (ec *ErrorsCollection) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "count":
		return len(ec.errors)
	}
	return nil
}

func (ec *ErrorsCollection) SetProperty(name string, value interface{}) {}

func (ec *ErrorsCollection) CallMethod(name string, args ...interface{}) interface{} {
	switch strings.ToLower(name) {
	case "item":
		if len(args) < 1 {
			return nil
		}
		idx := toInt(args[0])
		if idx < 0 || idx >= len(ec.errors) {
			return nil
		}
		return ec.errors[idx]
	case "clear":
		ec.Clear()
		return nil
	case "count":
		return len(ec.errors)
	}
	return nil
}

// --- ADODB.Connection ---

// ADODBConnection implements ADODB.Connection for database operations
type ADODBConnection struct {
	ConnectionString string
	State            int // 0 = closed, 1 = open
	Mode             int // Connection mode (typically 3 for read/write)
	db               *sql.DB
	ctx              *ExecutionContext
	dbDriver         string // Track which driver is in use
	Errors           *ErrorsCollection
	oleConnection    *ole.IDispatch // For Access databases via OLE
	oleInitialized   bool           // Track if COM was initialized for this connection
	threadLocked     bool           // Track if the OS thread is locked for COM usage
	tx               *sql.Tx        // Active transaction
}

// NewADODBConnection creates a new connection object
func NewADODBConnection(ctx *ExecutionContext) *ADODBConnection {
	conn := &ADODBConnection{
		State:  0,
		Mode:   3,
		ctx:    ctx,
		Errors: NewErrorsCollection(),
	}
	// Register for automatic cleanup
	if ctx != nil {
		ctx.RegisterManagedResource(conn)
	}
	return conn
}

func (c *ADODBConnection) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "connectionstring":
		return c.ConnectionString
	case "state":
		return c.State
	case "mode":
		return c.Mode
	case "errors":
		return c.Errors
	}
	return nil
}

func (c *ADODBConnection) SetProperty(name string, value interface{}) {
	switch strings.ToLower(name) {
	case "connectionstring":
		c.ConnectionString = fmt.Sprintf("%v", value)
		fmt.Printf("ADODB.Connection ConnectionString set to: %s\n", c.ConnectionString)
	case "mode":
		c.Mode = toInt(value)
	}
}

func (c *ADODBConnection) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	switch method {
	case "open":
		// Open([ConnectionString], [UserID], [Password], [Options])
		if len(args) > 0 {
			c.ConnectionString = fmt.Sprintf("%v", args[0])
		}
		if err := c.openDatabase(); err != nil {
			c.Errors.AddError(-1, err.Error(), "ADODB.Connection", "")
		}
		return nil

	case "close":
		if c.db != nil {
			c.db.Close()
			c.db = nil
		}
		if c.oleConnection != nil {
			// Safely close OLE connection
			defer func() {
				if r := recover(); r != nil {
					// Silently recover from any panic during OLE cleanup
				}
			}()
			oleutil.CallMethod(c.oleConnection, "Close")
			c.oleConnection.Release()
			c.oleConnection = nil
		}
		// Uninitialize COM and unlock thread if they were initialized for this connection
		if c.oleInitialized {
			ole.CoUninitialize()
			c.oleInitialized = false
		}
		if c.threadLocked {
			runtime.UnlockOSThread()
			c.threadLocked = false
		}
		c.State = 0
		return nil

	case "execute":
		// Execute(CommandText, [RecordsAffected], [Options])
		c.Errors.Clear()
		if len(args) < 1 {
			c.Errors.AddError(-1, "Invalid parameters", "ADODB.Connection", "")
			return nil
		}

		// Check if connection is open
		if c.db == nil && c.oleConnection == nil {
			c.Errors.AddError(-1, "Connection not open", "ADODB.Connection", "")
			return nil
		}

		cmdText := fmt.Sprintf("%v", args[0])
		params := []interface{}{}
		if len(args) >= 2 {
			if vbArr, ok := toVBArray(args[1]); ok {
				params = vbArr.Values
			} else if slice, ok := args[1].([]interface{}); ok {
				params = slice
			}
		}

		// Handle SQL driver connections
		if c.db != nil {
			if isQueryStatement(cmdText) {
				rs := NewADORecordset(c.ctx)
				if err := rs.lib.openRecordsetWithParams(cmdText, c, params); err != nil {
					c.Errors.AddError(-1, err.Error(), "ADODB.Connection", "")
					return nil
				}
				return rs
			}

			result, err := c.execStatement(cmdText, params)
			if err != nil {
				c.Errors.AddError(-1, err.Error(), "ADODB.Connection", "")
				return nil
			}
			affected, _ := result.RowsAffected()
			return int(affected)
		}

		// Handle OLE/Access connections - return Recordset
		if c.oleConnection != nil {
			result, err := oleutil.CallMethod(c.oleConnection, "Execute", cmdText)
			if err != nil {
				c.Errors.AddError(-1, err.Error(), "ADODB.Connection", "")
				return nil
			}

			if result != nil {
				oleRs := result.ToIDispatch()
				if oleRs != nil {
					return NewADOOLERecordset(NewADODBOLERecordset(oleRs, c.ctx))
				}
			}
			return nil
		}

		return nil

	case "begintrans":
		if c.db == nil {
			return nil
		}
		if c.tx != nil {
			return nil // Transaction already active
		}
		tx, err := c.db.Begin()
		if err != nil {
			c.Errors.AddError(-1, err.Error(), "ADODB.Connection", "")
			return nil
		}
		c.tx = tx
		return 1 // Level of transaction

	case "committrans":
		if c.tx == nil {
			return nil
		}
		if err := c.tx.Commit(); err != nil {
			c.Errors.AddError(-1, err.Error(), "ADODB.Connection", "")
		}
		c.tx = nil
		return nil

	case "rollbacktrans":
		if c.tx == nil {
			return nil
		}
		if err := c.tx.Rollback(); err != nil {
			c.Errors.AddError(-1, err.Error(), "ADODB.Connection", "")
		}
		c.tx = nil
		return nil
	}

	return nil
}

// openDatabase parses connection string and opens database
func (c *ADODBConnection) openDatabase() error {
	connStr := strings.TrimSpace(c.ConnectionString)
	if connStr == "" {
		return fmt.Errorf("connection string is empty")
	}

	// Check for Microsoft Access formats - use OLE method
	connStrLower := strings.ToLower(connStr)
	if (strings.Contains(connStrLower, "microsoft.jet.oledb") || strings.Contains(connStrLower, "microsoft.ace.oledb")) && runtime.GOOS == "windows" {
		return c.openAccessDatabase(connStr)
	}

	// Parse ODBC-style connection string for other drivers
	driver, dsn := parseConnectionString(connStr)

	if driver == "" || dsn == "" {
		return fmt.Errorf("unsupported or invalid connection string")
	}

	db, err := sql.Open(driver, dsn)
	if err != nil {
		return err
	}

	// Test connection
	if err := db.Ping(); err != nil {
		db.Close()
		return err
	}

	c.db = db
	c.dbDriver = driver
	c.State = 1
	return nil
}

func (c *ADODBConnection) execStatement(sqlText string, params []interface{}) (sql.Result, error) {
	if c.db == nil {
		return nil, fmt.Errorf("connection not open")
	}

	prepared := rewritePlaceholders(sqlText, c.dbDriver)
	if c.tx != nil {
		return c.tx.Exec(prepared, params...)
	}
	return c.db.Exec(prepared, params...)
}

func (c *ADODBConnection) queryRows(sqlText string, params []interface{}) (*sql.Rows, error) {
	if c.db == nil {
		return nil, fmt.Errorf("connection not open")
	}

	prepared := rewritePlaceholders(sqlText, c.dbDriver)
	if c.tx != nil {
		return c.tx.Query(prepared, params...)
	}
	return c.db.Query(prepared, params...)
}

func normalizeAccessProvider(connStr string) string {
	if strings.TrimSpace(connStr) == "" {
		return connStr
	}
	if GetCOMProviderMode() != "auto" {
		return connStr
	}

	provider := "microsoft.ace.oledb.12.0"
	if runtime.GOARCH == "386" {
		provider = "microsoft.jet.oledb.4.0"
	}

	parts := strings.Split(connStr, ";")
	found := false
	for i, part := range parts {
		if strings.TrimSpace(part) == "" {
			continue
		}
		kv := strings.SplitN(part, "=", 2)
		if len(kv) != 2 {
			continue
		}
		key := strings.TrimSpace(kv[0])
		if strings.EqualFold(key, "provider") {
			parts[i] = fmt.Sprintf("Provider=%s", provider)
			found = true
		}
	}
	if !found {
		parts = append([]string{fmt.Sprintf("Provider=%s", provider)}, parts...)
	}
	return strings.Join(parts, ";")
}

// openAccessDatabase opens an Access database using OLE/OLEDB
func (c *ADODBConnection) openAccessDatabase(connStr string) error {
	// Only supported on Windows
	if runtime.GOOS != "windows" {
		fmt.Println("Warning: Direct Access database support is only available on Windows. Please use a different database system for cross-platform compatibility.")
		return fmt.Errorf("access database support is only available on windows")
	}

	connStr = normalizeAccessProvider(connStr)

	// Ensure COM calls stay on a single OS thread for the lifetime of this connection
	runtime.LockOSThread()
	c.threadLocked = true

	// Initialize COM for this thread - DO NOT uninitialize here, it will be done in Close()
	if err := ole.CoInitialize(0); err != nil {
		fmt.Printf("Warning: Failed to initialize COM for Access database. Error details: %v\n", err)
		runtime.UnlockOSThread()
		c.threadLocked = false
		return err
	}
	c.oleInitialized = true

	unknown, err := oleutil.CreateObject("ADODB.Connection")
	if err != nil {
		fmt.Printf("Warning: ADODB.Connection COM object cannot be created. Error: %v. Make sure you have Windows COM support and OLEDB drivers installed.\n", err)
		ole.CoUninitialize()
		c.oleInitialized = false
		runtime.UnlockOSThread()
		c.threadLocked = false
		return err
	}

	connection, err := unknown.QueryInterface(ole.IID_IDispatch)
	// Release unknown after getting IDispatch - we don't need it anymore
	unknown.Release()

	if err != nil {
		fmt.Println("Warning: Cannot query ADODB.Connection interface. COM support may not be available.")
		ole.CoUninitialize()
		c.oleInitialized = false
		runtime.UnlockOSThread()
		c.threadLocked = false
		return err
	}

	// Open the connection
	_, err = oleutil.CallMethod(connection, "Open", connStr)
	if err != nil {
		fmt.Printf("Warning: Cannot open Access database. Error details: %v\n", err)
		fmt.Printf("Connection string: %s\n", connStr)
		connection.Release()
		ole.CoUninitialize()
		c.oleInitialized = false
		runtime.UnlockOSThread()
		c.threadLocked = false
		return err
	}

	// Access database opened successfully
	// Store the OLE connection object - it needs to stay alive for queries
	c.oleConnection = connection
	c.State = 1
	return nil
}

// parseConnectionString converts ODBC connection string to Go SQL driver and DSN
func parseConnectionString(connStr string) (driver string, dsn string) {
	trimmed := strings.TrimSpace(connStr)
	connStrLower := strings.ToLower(trimmed)

	// Handle SQLite formats
	if strings.HasPrefix(connStrLower, "sqlite:") {
		driver = "sqlite"
		dbPath := strings.TrimPrefix(trimmed, trimmed[:len("sqlite:")])
		dsn = strings.TrimSpace(dbPath)
		if dsn == "" {
			dsn = ":memory:"
		}
		return
	}

	// Access databases should be handled via OLE (see openAccessDatabase)
	// If we reach here, Access is not supported on this platform
	if strings.Contains(connStrLower, "microsoft.jet.oledb") || strings.Contains(connStrLower, "microsoft.ace.oledb") {
		fmt.Println("Warning: Direct Access database support is only available on Windows. Please use a different database system for cross-platform compatibility.")
		return
	}

	// Parse ODBC-style: Driver={...};Server=...;Database=...;UID=...;PWD=...
	params := make(map[string]string)
	parts := strings.Split(trimmed, ";")
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if idx := strings.Index(part, "="); idx > 0 {
			key := strings.ToLower(strings.TrimSpace(part[:idx]))
			val := strings.TrimSpace(part[idx+1:])
			// Remove curly braces if present
			val = strings.Trim(val, "{}")
			params[key] = val
		}
	}

	driverStr := strings.ToLower(params["driver"])
	providerStr := strings.ToLower(params["provider"])

	server := firstNonEmpty(params["server"], params["data source"], params["datasource"], params["host"], params["address"], params["addr"])
	database := firstNonEmpty(params["database"], params["initial catalog"], params["dbname"], params["data source name"])
	uid := firstNonEmpty(params["uid"], params["user id"], params["user"], params["username"])
	pwd := firstNonEmpty(params["pwd"], params["password"], params["pass"])

	// SQLite (ODBC-style)
	if strings.Contains(driverStr, "sqlite") {
		driver = "sqlite"
		dsn = firstNonEmpty(params["data source"], params["datasource"], params["database"])
		if dsn == "" {
			dsn = ":memory:"
		}
		return
	}

	// MySQL
	if strings.Contains(driverStr, "mysql") {
		driver = "mysql"
		port := params["port"]
		if port == "" {
			port = "3306"
		}
		dsn = fmt.Sprintf("%s:%s@tcp(%s:%s)/%s?parseTime=true", uid, pwd, server, port, database)
		return
	}

	// PostgreSQL
	if strings.Contains(driverStr, "postgresql") || strings.Contains(driverStr, "postgres") {
		driver = "postgres"
		port := params["port"]
		if port == "" {
			port = "5432"
		}
		dsn = fmt.Sprintf("user=%s password=%s host=%s port=%s dbname=%s sslmode=disable", uid, pwd, server, port, database)
		return
	}

	// MS SQL Server
	if strings.Contains(driverStr, "sql server") || strings.Contains(driverStr, "mssql") || strings.Contains(providerStr, "sqloledb") || strings.Contains(providerStr, "sqlncli") {
		driver = "mssql"
		port := params["port"]
		if port == "" {
			port = "1433"
		}
		dsn = fmt.Sprintf("server=%s;user id=%s;password=%s;database=%s;port=%s", server, uid, pwd, database, port)
		return
	}

	return
}

func firstNonEmpty(values ...string) string {
	for _, value := range values {
		value = strings.TrimSpace(value)
		if value != "" {
			return value
		}
	}
	return ""
}

func normalizeExecuteParams(args []interface{}) []interface{} {
	if len(args) == 0 {
		return nil
	}
	if len(args) == 1 {
		if vbArr, ok := toVBArray(args[0]); ok {
			return vbArr.Values
		}
		if slice, ok := args[0].([]interface{}); ok {
			return slice
		}
	}
	return args
}

func isQueryStatement(sqlText string) bool {
	trimmed := strings.TrimSpace(strings.ToLower(sqlText))
	if trimmed == "" {
		return false
	}
	for _, prefix := range []string{"select", "with", "show", "pragma", "describe", "exec", "execute", "call"} {
		if strings.HasPrefix(trimmed, prefix) {
			return true
		}
	}
	return false
}

func rewritePlaceholders(sqlText string, driver string) string {
	if driver != "postgres" && driver != "mssql" {
		return sqlText
	}

	var out strings.Builder
	out.Grow(len(sqlText) + 16)

	inSingle := false
	inDouble := false
	paramIndex := 0

	for i := 0; i < len(sqlText); i++ {
		ch := sqlText[i]

		if ch == '\'' && !inDouble {
			if inSingle && i+1 < len(sqlText) && sqlText[i+1] == '\'' {
				out.WriteByte(ch)
				i++
				out.WriteByte(sqlText[i])
				continue
			}
			inSingle = !inSingle
			out.WriteByte(ch)
			continue
		}
		if ch == '"' && !inSingle {
			if inDouble && i+1 < len(sqlText) && sqlText[i+1] == '"' {
				out.WriteByte(ch)
				i++
				out.WriteByte(sqlText[i])
				continue
			}
			inDouble = !inDouble
			out.WriteByte(ch)
			continue
		}

		if ch == '?' && !inSingle && !inDouble {
			paramIndex++
			if driver == "postgres" {
				out.WriteString(fmt.Sprintf("$%d", paramIndex))
			} else {
				out.WriteString(fmt.Sprintf("@p%d", paramIndex))
			}
			continue
		}

		out.WriteByte(ch)
	}

	return out.String()
}

// --- ADODB.Recordset ---

// Field represents a database field with Name and Value
// Implements asp.ASPObject interface for VBScript compatibility
type Field struct {
	Name  string
	Value interface{}
}

func (f *Field) GetName() string {
	return "Field"
}

func (f *Field) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "name":
		return f.Name
	case "value":
		return f.Value
	}
	return nil
}

func (f *Field) SetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "value":
		f.Value = value
	}
	return nil
}

func (f *Field) CallMethod(name string, args ...interface{}) (interface{}, error) {
	// Default method returns value
	if strings.ToLower(name) == "" {
		return f.Value, nil
	}
	return nil, nil
}

func (f *Field) String() string {
	if f == nil {
		return ""
	}
	return fmt.Sprintf("%v", f.Value)
}

// OLEFieldProxy provides field access for OLE recordsets without holding COM references.
// It resolves the field value dynamically via the parent recordset.
type OLEFieldProxy struct {
	recordset *ADODBOLERecordset
	name      string
}

func (f *OLEFieldProxy) GetName() string {
	return "Field"
}

func (f *OLEFieldProxy) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "name":
		return f.name
	case "value":
		if f.recordset == nil {
			return nil
		}
		value, _ := f.recordset.getFieldValue(f.name)
		return value
	}
	return nil
}

func (f *OLEFieldProxy) SetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "value":
		if f.recordset != nil {
			_ = f.recordset.setFieldValue(f.name, value)
		}
	}
	return nil
}

func (f *OLEFieldProxy) CallMethod(name string, args ...interface{}) (interface{}, error) {
	if strings.ToLower(name) == "" {
		return f.GetProperty("value"), nil
	}
	return nil, nil
}

func (f *OLEFieldProxy) String() string {
	if f == nil {
		return ""
	}
	return fmt.Sprintf("%v", f.GetProperty("value"))
}

// FieldsCollection holds a collection of Field objects
// Implements asp.ASPObject interface for VBScript compatibility
type FieldsCollection struct {
	fields []*Field
	data   map[string]interface{}
}

func NewFieldsCollection() *FieldsCollection {
	return &FieldsCollection{
		fields: make([]*Field, 0),
		data:   make(map[string]interface{}),
	}
}

func (fc *FieldsCollection) GetName() string {
	return "Fields"
}

func (fc *FieldsCollection) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "count":
		return len(fc.fields)
	case "item":
		// Item requires an argument, return nil to indicate method call needed
		return nil
	default:
		// Try to access as a field name directly
		if val, ok := fc.data[strings.ToLower(name)]; ok {
			return val
		}
	}
	return nil
}

func (fc *FieldsCollection) SetProperty(name string, value interface{}) error {
	return nil
}

func (fc *FieldsCollection) CallMethod(name string, args ...interface{}) (interface{}, error) {
	nameLower := strings.ToLower(name)
	// Default method is "Item" - handle empty name, "item", or any direct subscript access
	if nameLower == "" || nameLower == "item" {
		if len(args) < 1 {
			return nil, nil
		}
		// Item can be by index (int) or by name (string)
		// Returns the Field object (not just the value) so .name and .value can be accessed
		if idx, ok := args[0].(int); ok && idx >= 0 && idx < len(fc.fields) {
			return fc.fields[idx], nil
		}
		// Also try int32 (common from VBScript)
		if idx, ok := args[0].(int32); ok && int(idx) >= 0 && int(idx) < len(fc.fields) {
			return fc.fields[int(idx)], nil
		}
		// Also try int64
		if idx, ok := args[0].(int64); ok && int(idx) >= 0 && int(idx) < len(fc.fields) {
			return fc.fields[int(idx)], nil
		}
		// Also try float64 (common from JSON/VBScript integer literals)
		if idx, ok := args[0].(float64); ok && int(idx) >= 0 && int(idx) < len(fc.fields) {
			return fc.fields[int(idx)], nil
		}
		// Try by name
		key := strings.ToLower(fmt.Sprintf("%v", args[0]))
		for _, field := range fc.fields {
			if strings.ToLower(field.Name) == key {
				return field, nil
			}
		}
		return nil, nil
	}
	switch nameLower {
	case "count":
		return len(fc.fields), nil
	}
	return nil, nil
}

// Enumeration returns all Field objects for For Each iteration
func (fc *FieldsCollection) Enumeration() []interface{} {
	result := make([]interface{}, len(fc.fields))
	for i, field := range fc.fields {
		result[i] = field
	}
	return result
}

// ADODBRecordset simulates ADODB.Recordset for result sets
type ADODBRecordset struct {
	EOF               bool
	BOF               bool
	RecordCount       int
	State             int // 0 = closed, 1 = open
	CurrentRow        int
	Fields            *FieldsCollection
	rows              *sql.Rows
	db                *sql.DB
	dbConn            *ADODBConnection // Reference to parent connection wrapper for Tx support
	columns           []string
	currentData       map[string]interface{}
	allData           []map[string]interface{}
	originalData      []map[string]interface{}
	newData           map[string]interface{} // For AddNew
	ctx               *ExecutionContext
	PageSize          int         // Number of records per page
	AbsolutePage      int         // Current page number (1-based)
	PageCount         int         // Total number of pages
	SortField         string      // Field name for sorting
	SortOrder         string      // ASC or DESC
	FilterCriteria    string      // Filter expression
	filteredIndices   []int       // Indices of filtered records
	isFiltered        bool        // Whether filter is active
	ActiveConnection  interface{} // Stored connection for later use
	CursorType        int         // Cursor type
	LockType          int         // Lock type
	CursorLocation    int         // Cursor location (2=Server, 3=Client)
	CursorTypeSet     bool        // Track explicit CursorType assignment
	LockTypeSet       bool        // Track explicit LockType assignment
	CursorLocationSet bool        // Track explicit CursorLocation assignment
	Source            string      // Source SQL or Table Name
	oleRecordset      *ADODBOLERecordset
}

// NewADODBRecordset creates a new recordset
func NewADODBRecordset(ctx *ExecutionContext) *ADODBRecordset {
	return &ADODBRecordset{
		EOF:            true,
		BOF:            true,
		State:          0,
		CurrentRow:     -1,
		Fields:         NewFieldsCollection(),
		allData:        make([]map[string]interface{}, 0),
		originalData:   make([]map[string]interface{}, 0),
		ctx:            ctx,
		PageSize:       10,
		AbsolutePage:   1,
		PageCount:      0,
		SortOrder:      "ASC",
		CursorLocation: 2,
	}
}

func (rs *ADODBRecordset) GetProperty(name string) interface{} {
	if rs.oleRecordset != nil {
		return rs.oleRecordset.GetProperty(name)
	}
	nameLower := strings.ToLower(name)
	switch nameLower {
	case "eof":
		return rs.EOF
	case "bof":
		return rs.BOF
	case "recordcount":
		if rs.isFiltered {
			return len(rs.filteredIndices)
		}
		return rs.RecordCount
	case "state":
		return rs.State
	case "activeconnection":
		return rs.ActiveConnection
	case "source":
		return rs.Source
	case "cursorlocation":
		return rs.CursorLocation
	case "fields":
		return rs.Fields
	case "absoluteposition":
		if rs.CurrentRow >= 0 {
			return rs.CurrentRow + 1
		}
		return 0
	case "pagesize":
		return rs.PageSize
	case "absolutepage":
		return rs.AbsolutePage
	case "pagecount":
		rs.calculatePageCount()
		return rs.PageCount
	case "sort":
		return rs.SortField
	case "filter":
		return rs.FilterCriteria
	case "getrows":
		// getRows called without parentheses - call as method with no args
		return rs.CallMethod("getrows")
	}
	return nil
}

func (rs *ADODBRecordset) SetProperty(name string, value interface{}) {
	nameLower := strings.ToLower(name)
	if nameLower == "activeconnection" {
		rs.ActiveConnection = value
		return
	}
	if rs.oleRecordset != nil {
		rs.oleRecordset.SetProperty(name, value)
		return
	}
	switch nameLower {
	case "activeconnection":
		rs.ActiveConnection = value
	case "cursortype":
		rs.CursorType = toInt(value)
		rs.CursorTypeSet = true
	case "locktype":
		rs.LockType = toInt(value)
		rs.LockTypeSet = true
	case "cursorlocation":
		rs.CursorLocation = toInt(value)
		rs.CursorLocationSet = true
	case "source":
		rs.Source = fmt.Sprintf("%v", value)
	case "absoluteposition":
		pos := toInt(value)
		if pos > 0 {
			rs.setCurrentRowByIndex(pos - 1)
		}
	case "pagesize":
		v := toInt(value)
		if v > 0 {
			rs.PageSize = v
			rs.calculatePageCount()
		}
	case "absolutepage":
		v := toInt(value)
		if v > 0 {
			rs.AbsolutePage = v
			rs.moveToAbsolutePage()
		}
	case "sort":
		if v, ok := value.(string); ok {
			rs.applySorting(v)
		}
	case "filter":
		if v, ok := value.(string); ok {
			if v == "" {
				rs.clearFilter()
			} else {
				rs.applyFilter(v)
			}
		} else if value == nil || value == 0 {
			// Clear filter
			rs.clearFilter()
		}
	default:
		// Handle field assignment: rs("field") = value
		// If in AddNew mode
		if rs.newData != nil {
			rs.newData[nameLower] = value
			return
		}

		// If editing current record
		if rs.CurrentRow >= 0 && rs.CurrentRow < len(rs.allData) {
			// Update currentData (in-memory)
			if rs.currentData == nil {
				rs.currentData = make(map[string]interface{})
			}
			rs.currentData[nameLower] = value
			rs.allData[rs.CurrentRow][nameLower] = value

			// Update Fields collection to reflect change immediately
			found := false
			for _, f := range rs.Fields.fields {
				if strings.ToLower(f.Name) == nameLower {
					f.Value = value
					found = true
					break
				}
			}
			if found {
				rs.Fields.data[nameLower] = value
			}
		}
	}
}

func (rs *ADODBRecordset) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)
	if rs.oleRecordset != nil && method != "open" {
		if method == "close" {
			rs.oleRecordset.CallMethod("close")
			rs.oleRecordset = nil
			rs.State = 0
			rs.EOF = true
			rs.BOF = true
			return nil
		}
		return rs.oleRecordset.CallMethod(name, args...)
	}

	// Handle default method (empty name) - returns field VALUE, not Field object
	if method == "" && len(args) > 0 {
		// Get the field by name or index and return its value
		field, _ := rs.Fields.CallMethod("item", args...)
		if f, ok := field.(*Field); ok {
			return f.Value
		}
		return field // In case it's already a value
	}

	switch method {
	case "fields":
		// rs.fields(i) - delegate to FieldsCollection
		// If called with args, return the field at that index/name
		// If called without args, return the Fields collection itself
		if len(args) > 0 {
			result, _ := rs.Fields.CallMethod("item", args...)
			return result
		}
		return rs.Fields

	case "open":
		// Open(Source, [ActiveConnection], [CursorType], [LockType], [Options])
		if len(args) < 1 {
			return nil
		}
		sql := fmt.Sprintf("%v", args[0])

		// Try to get connection from args[1] or from stored ActiveConnection
		var conn *ADODBConnection

		if len(args) >= 2 && args[1] != nil {
			// Connection provided as argument
			if c, ok := args[1].(*ADODBConnection); ok {
				conn = c
			} else if cWrapper, ok := args[1].(*ADOConnection); ok {
				conn = cWrapper.lib
			}
		}

		// If no connection in args, try stored ActiveConnection
		if conn == nil && rs.ActiveConnection != nil {
			if c, ok := rs.ActiveConnection.(*ADODBConnection); ok {
				conn = c
			} else if cWrapper, ok := rs.ActiveConnection.(*ADOConnection); ok {
				conn = cWrapper.lib
			}
		}

		if conn == nil {
			return nil
		}
		if conn.oleConnection != nil {
			return rs.openOLERecordset(args, conn)
		}

		return rs.openRecordset(sql, conn)

	case "close":
		if rs.rows != nil {
			rs.rows.Close()
		}
		rs.rows = nil
		rs.currentData = nil
		rs.allData = nil
		rs.originalData = nil
		rs.columns = nil
		rs.State = 0
		rs.EOF = true
		rs.BOF = true
		return nil

	case "movenext":
		if rs.oleRecordset != nil {
			return rs.oleRecordset.CallMethod("movenext")
		}
		if rs.EOF {
			return nil
		}
		if rs.isFiltered {
			// Find next index in filtered list
			currentFilteredIdx := rs.findCurrentFilteredIndex()
			if currentFilteredIdx >= 0 && currentFilteredIdx+1 < len(rs.filteredIndices) {
				rs.CurrentRow = rs.filteredIndices[currentFilteredIdx+1]
				rs.fetchCurrentRow()
				rs.BOF = false
			} else {
				rs.EOF = true
			}
		} else {
			if rs.CurrentRow+1 < len(rs.allData) {
				rs.CurrentRow++
				rs.fetchCurrentRow()
				rs.BOF = false
			} else {
				rs.EOF = true
			}
		}
		return nil

	case "movefirst":
		if rs.oleRecordset != nil {
			return rs.oleRecordset.CallMethod("movefirst")
		}
		if rs.isFiltered {
			if len(rs.filteredIndices) > 0 {
				rs.CurrentRow = rs.filteredIndices[0]
				rs.currentData = rs.allData[rs.CurrentRow]
				rs.updateFieldsCollection()
				rs.BOF = false
				rs.EOF = false
			} else {
				rs.EOF = true
				rs.BOF = true
			}
		} else {
			if len(rs.allData) > 0 {
				rs.CurrentRow = 0
				rs.currentData = rs.allData[0]
				rs.updateFieldsCollection()
				rs.BOF = false
				rs.EOF = false
			} else {
				rs.EOF = true
				rs.BOF = true
			}
		}
		return nil

	case "movelast":
		if rs.oleRecordset != nil {
			return rs.oleRecordset.CallMethod("movelast")
		}
		if rs.isFiltered {
			if len(rs.filteredIndices) == 0 {
				rs.EOF = true
				rs.BOF = true
				return nil
			}
			rs.CurrentRow = rs.filteredIndices[len(rs.filteredIndices)-1]
			rs.currentData = rs.allData[rs.CurrentRow]
			rs.updateFieldsCollection()
			rs.EOF = false
			rs.BOF = false
		} else {
			if len(rs.allData) == 0 {
				rs.EOF = true
				rs.BOF = true
				return nil
			}
			rs.CurrentRow = len(rs.allData) - 1
			rs.currentData = rs.allData[rs.CurrentRow]
			rs.updateFieldsCollection()
			rs.EOF = false
			rs.BOF = false
		}
		return nil

	case "moveprevious":
		if rs.isFiltered {
			currentFilteredIdx := rs.findCurrentFilteredIndex()
			if currentFilteredIdx > 0 {
				rs.CurrentRow = rs.filteredIndices[currentFilteredIdx-1]
				rs.currentData = rs.allData[rs.CurrentRow]
				rs.updateFieldsCollection()
				rs.EOF = false
				rs.BOF = false
			} else if currentFilteredIdx == 0 {
				rs.BOF = true
			}
		} else {
			if rs.CurrentRow > 0 {
				rs.CurrentRow--
				rs.currentData = rs.allData[rs.CurrentRow]
				rs.updateFieldsCollection()
				rs.EOF = false
				rs.BOF = false
			} else if rs.CurrentRow == 0 {
				rs.BOF = true
			}
		}
		return nil

	case "move":
		if len(args) < 1 {
			return nil
		}
		offset := toInt(args[0])
		startProvided := len(args) >= 2
		startIndex := -1
		if startProvided {
			startIndex = toInt(args[1]) - 1
		}

		if rs.isFiltered {
			currentFilteredIdx := rs.findCurrentFilteredIndex()
			base := currentFilteredIdx
			if startProvided {
				base = startIndex
			}
			rs.setCurrentRowByIndex(base + offset)
			return nil
		}

		base := rs.CurrentRow
		if startProvided {
			base = startIndex
		}
		rs.setCurrentRowByIndex(base + offset)
		return nil

	case "addnew":
		rs.newData = make(map[string]interface{})
		if len(args) >= 2 {
			fields := toStringSlice(args[0])
			values := toInterfaceSlice(args[1])
			for i, field := range fields {
				if i < len(values) {
					rs.newData[strings.ToLower(field)] = values[i]
				}
			}
		}
		return nil

	case "update":
		if len(args) >= 2 {
			rs.SetProperty(fmt.Sprintf("%v", args[0]), args[1])
		}
		// Handle DB persistence
		if rs.db != nil && rs.Source != "" {
			rs.performSQLUpdate()
		}

		// Add new row to allData
		if rs.newData != nil {
			rs.allData = append(rs.allData, rs.newData)
			rs.originalData = append(rs.originalData, cloneRow(rs.newData))
			rs.RecordCount = len(rs.allData)
			rs.CurrentRow = len(rs.allData) - 1
			rs.currentData = rs.allData[rs.CurrentRow]
			rs.updateFieldsCollection()
			rs.newData = nil
		}
		return nil

	case "delete":
		if rs.db != nil && rs.Source != "" {
			rs.performSQLDelete()
		}
		if rs.CurrentRow >= 0 && rs.CurrentRow < len(rs.allData) {
			rs.allData = append(rs.allData[:rs.CurrentRow], rs.allData[rs.CurrentRow+1:]...)
			if rs.CurrentRow >= 0 && rs.CurrentRow < len(rs.originalData) {
				rs.originalData = append(rs.originalData[:rs.CurrentRow], rs.originalData[rs.CurrentRow+1:]...)
			}
			rs.RecordCount = len(rs.allData)
			if rs.RecordCount == 0 {
				rs.CurrentRow = -1
				rs.EOF = true
				rs.BOF = true
			} else if rs.CurrentRow >= rs.RecordCount {
				rs.CurrentRow = rs.RecordCount - 1
				rs.fetchCurrentRow()
			} else {
				rs.fetchCurrentRow()
			}
		}
		return nil

	case "cancelupdate":
		rs.newData = nil
		if rs.CurrentRow >= 0 && rs.CurrentRow < len(rs.originalData) {
			rs.allData[rs.CurrentRow] = cloneRow(rs.originalData[rs.CurrentRow])
			rs.currentData = rs.allData[rs.CurrentRow]
			rs.updateFieldsCollection()
		}
		return nil

	case "item", "collect":
		// rs.Item("FieldName") or rs(0) style access - returns value
		if len(args) > 0 {
			field, _ := rs.Fields.CallMethod("item", args...)
			if f, ok := field.(*Field); ok {
				return f.Value
			}
			return field
		}
		return nil

	case "getrows":
		// GetRows([Rows], [Start], [Fields])
		// Returns 2D array with remaining records from current position
		return rs.getRows(args)

	case "find":
		// Find(Criteria, [SkipRecords], [SearchDirection], [Start])
		if len(args) < 1 {
			return nil
		}
		criteria := fmt.Sprintf("%v", args[0])
		rs.findRecord(criteria)
		return nil

	case "supports":
		// Supports(CursorOptions) - checks if provider supports a feature
		if len(args) < 1 {
			return false
		}
		option, ok := args[0].(int)
		if !ok {
			return false
		}
		return rs.supportsFeature(option)

	default:
		// Fallback for default property access (rs("field"))
		if len(args) > 0 {
			field, _ := rs.Fields.CallMethod("item", args...)
			if f, ok := field.(*Field); ok {
				return f.Value
			}
			return field
		}
	}

	return nil
}

func (rs *ADODBRecordset) openRecordset(sqlStr string, conn *ADODBConnection) interface{} {
	_ = rs.openRecordsetWithParams(sqlStr, conn, nil)
	return nil
}

func (rs *ADODBRecordset) openRecordsetWithParams(sqlStr string, conn *ADODBConnection, params []interface{}) error {
	rs.Source = sqlStr

	// Use SQL driver connection
	if conn.db == nil {
		return fmt.Errorf("connection not open")
	}

	rs.db = conn.db
	rs.dbConn = conn

	rows, err := conn.queryRows(sqlStr, params)
	if err != nil {
		return err
	}
	defer rows.Close()

	rs.rows = rows
	rs.State = 1
	rs.EOF = false
	rs.BOF = true
	rs.CurrentRow = -1
	rs.columns = nil
	rs.allData = make([]map[string]interface{}, 0)
	rs.originalData = make([]map[string]interface{}, 0)

	// Get column names
	cols, err := rows.Columns()
	if err != nil {
		return err
	}
	rs.columns = cols

	// Fetch all rows into memory for random access
	for rows.Next() {
		values := make([]interface{}, len(cols))
		valuePtrs := make([]interface{}, len(cols))
		for i := range cols {
			valuePtrs[i] = &values[i]
		}

		if err := rows.Scan(valuePtrs...); err != nil {
			continue
		}

		row := make(map[string]interface{})
		for i, col := range cols {
			row[strings.ToLower(col)] = values[i]
		}
		rs.allData = append(rs.allData, row)
		rs.originalData = append(rs.originalData, cloneRow(row))
	}

	rs.RecordCount = len(rs.allData)

	if len(rs.allData) > 0 {
		rs.CurrentRow = 0
		rs.currentData = rs.allData[0]
		rs.BOF = false
		rs.EOF = false
		rs.updateFieldsCollection()
	} else {
		rs.EOF = true
		rs.BOF = true
	}

	return nil
}

func (rs *ADODBRecordset) openOLERecordset(args []interface{}, conn *ADODBConnection) interface{} {
	if conn == nil || conn.oleConnection == nil {
		return nil
	}

	if rs.oleRecordset != nil {
		rs.oleRecordset.CallMethod("close")
		rs.oleRecordset = nil
	}

	unknown, err := oleutil.CreateObject("ADODB.Recordset")
	if err != nil {
		return nil
	}

	disp, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		unknown.Release()
		return nil
	}

	// DO NOT defer unknown.Release() - it's owned by the ADODBOLERecordset now

	// Apply CursorLocation property before Open
	cursorLocation := rs.CursorLocation
	if !rs.CursorLocationSet && !rs.CursorTypeSet && !rs.LockTypeSet && cursorLocation == 2 {
		cursorLocation = 3
	}
	if cursorLocation > 0 {
		oleutil.PutProperty(disp, "CursorLocation", toInt32Variant(cursorLocation))
	}

	openArgs := make([]interface{}, 0, 5)
	if len(args) > 0 {
		openArgs = append(openArgs, args[0])
	}
	if len(args) >= 2 {
		switch args[1].(type) {
		case *ADODBConnection, *ADOConnection:
			openArgs = append(openArgs, conn.oleConnection)
		default:
			openArgs = append(openArgs, args[1])
		}
	} else {
		openArgs = append(openArgs, conn.oleConnection)
	}

	// CursorType
	var cursorType interface{}
	if len(args) >= 3 {
		cursorType = toInt32Variant(args[2])
	} else if rs.CursorType > 0 {
		cursorType = toInt32Variant(rs.CursorType)
	} else if !rs.CursorTypeSet {
		// Default to Keyset (1) for Access to allow updates
		cursorType = int32(1)
	} else {
		// Default to ForwardOnly (0) if not specified
		cursorType = int32(0)
	}
	if toInt(cursorLocation) == 3 {
		if ct, ok := cursorType.(int32); ok && (ct == 1 || ct == 2) {
			cursorType = int32(3)
		}
	}
	openArgs = append(openArgs, cursorType)

	// LockType
	if len(args) >= 4 {
		openArgs = append(openArgs, toInt32Variant(args[3]))
	} else if rs.LockType > 0 {
		openArgs = append(openArgs, toInt32Variant(rs.LockType))
	} else if !rs.LockTypeSet {
		// Default to Optimistic (3) for Access to allow updates
		openArgs = append(openArgs, int32(3))
	} else {
		// Default to ReadOnly (1)
		openArgs = append(openArgs, int32(1))
	}

	if len(args) >= 5 {
		openArgs = append(openArgs, args[4])
	}

	if _, err = oleutil.CallMethod(disp, "Open", openArgs...); err != nil {
		disp.Release()
		unknown.Release()
		return nil
	}

	rs.oleRecordset = NewADODBOLERecordset(disp, rs.ctx)
	rs.oleRecordset.activeConnection = conn.oleConnection
	rs.State = 1
	rs.EOF = false
	rs.BOF = false
	return nil
}

func toInt32Variant(value interface{}) interface{} {
	switch v := value.(type) {
	case int:
		return int32(v)
	case int32:
		return v
	case int64:
		return int32(v)
	case float64:
		return int32(v)
	case float32:
		return int32(v)
	default:
		return value
	}
}

func (rs *ADODBRecordset) fetchCurrentRow() {
	if rs.CurrentRow >= 0 && rs.CurrentRow < len(rs.allData) {
		rs.currentData = rs.allData[rs.CurrentRow]
		rs.updateFieldsCollection()
	}
}

func (rs *ADODBRecordset) updateFieldsCollection() {
	rs.Fields.fields = make([]*Field, 0)
	rs.Fields.data = make(map[string]interface{})

	if rs.currentData != nil {
		// Use column order if available to maintain consistent field ordering
		if len(rs.columns) > 0 {
			for _, col := range rs.columns {
				colLower := strings.ToLower(col)
				value := rs.currentData[colLower]
				rs.Fields.fields = append(rs.Fields.fields, &Field{Name: col, Value: value})
				rs.Fields.data[colLower] = value
			}
		} else {
			// Fallback to map iteration (unordered)
			for name, value := range rs.currentData {
				rs.Fields.fields = append(rs.Fields.fields, &Field{Name: name, Value: value})
				rs.Fields.data[strings.ToLower(name)] = value
			}
		}
	}
}

func cloneRow(row map[string]interface{}) map[string]interface{} {
	if row == nil {
		return nil
	}
	clone := make(map[string]interface{}, len(row))
	for key, value := range row {
		clone[key] = value
	}
	return clone
}

func toInterfaceSlice(value interface{}) []interface{} {
	if value == nil {
		return nil
	}
	if vbArr, ok := toVBArray(value); ok {
		return vbArr.Values
	}
	if slice, ok := value.([]interface{}); ok {
		return slice
	}
	return []interface{}{value}
}

func toStringSlice(value interface{}) []string {
	values := toInterfaceSlice(value)
	if len(values) == 0 {
		return nil
	}
	result := make([]string, 0, len(values))
	for _, item := range values {
		result = append(result, fmt.Sprintf("%v", item))
	}
	return result
}

// populateRecordsetFromOLE reads all rows from an OLE recordset into an ADODBRecordset
// and releases the COM recordset. Returns true on success.
func populateRecordsetFromOLE(oleRs *ole.IDispatch, rs *ADODBRecordset) bool {
	if oleRs == nil || rs == nil {
		return false
	}

	defer func() {
		// Ensure COM resources are released even on early return
		oleRs.Release()
	}()

	rs.State = 1
	rs.allData = make([]map[string]interface{}, 0)

	fieldsResult, err := oleutil.GetProperty(oleRs, "Fields")
	if err != nil {
		return false
	}

	fieldsObj := fieldsResult.ToIDispatch()
	if fieldsObj == nil {
		return false
	}
	defer fieldsObj.Release()

	countResult, _ := oleutil.GetProperty(fieldsObj, "Count")
	fieldCount := int(countResult.Val)
	if fieldCount < 0 {
		fieldCount = 0
	}

	rs.columns = make([]string, fieldCount)
	for i := 0; i < fieldCount; i++ {
		itemResult, _ := oleutil.GetProperty(fieldsObj, "Item", i)
		field := itemResult.ToIDispatch()
		if field == nil {
			continue
		}
		nameResult, _ := oleutil.GetProperty(field, "Name")
		rs.columns[i] = fmt.Sprintf("%v", nameResult.Value())
		field.Release()
	}

	for {
		eofResult, eofErr := oleutil.GetProperty(oleRs, "EOF")
		if eofErr != nil {
			break
		}
		isEOF := false
		switch v := eofResult.Value().(type) {
		case bool:
			isEOF = v
		case int, int32, int64:
			isEOF = fmt.Sprintf("%v", v) != "0"
		default:
			if eofResult.Val != 0 {
				isEOF = true
			}
		}
		if isEOF {
			break
		}

		row := make(map[string]interface{})
		for i := 0; i < fieldCount; i++ {
			itemResult, _ := oleutil.GetProperty(fieldsObj, "Item", i)
			field := itemResult.ToIDispatch()
			if field == nil {
				continue
			}
			valueResult, err := oleutil.GetProperty(field, "Value")
			colName := ""
			if i < len(rs.columns) {
				colName = strings.ToLower(rs.columns[i])
			}
			if err != nil {
				row[colName] = nil
			} else {
				row[colName] = valueResult.Value()
			}
			field.Release()
		}
		rs.allData = append(rs.allData, row)

		oleutil.CallMethod(oleRs, "MoveNext")
	}

	rs.RecordCount = len(rs.allData)

	if len(rs.allData) > 0 {
		rs.CurrentRow = 0
		rs.currentData = rs.allData[0]
		rs.BOF = false
		rs.EOF = false
		rs.updateFieldsCollection()
	} else {
		rs.EOF = true
		rs.BOF = true
	}

	return true
}

// getRows returns a 2D array containing all records from current position to end
// NOTE: In classic ADO, GetRows advances the cursor to EOF. However, to support
// scenarios where the expression is evaluated multiple times (e.g., when passed
// directly as a function argument), we reset to the first record if EOF is true
// but data is available.
func (rs *ADODBRecordset) getRows(args []interface{}) interface{} {
	if len(rs.allData) == 0 {
		return [][]interface{}{}
	}

	// If EOF but we have data, reset to beginning (defensive handling)
	if rs.EOF && len(rs.allData) > 0 {
		rs.CurrentRow = 0
		rs.EOF = false
		rs.BOF = false
	}

	// Determine how many rows to fetch
	maxRows := -1 // -1 means all remaining rows
	if len(args) > 0 {
		if v, ok := args[0].(int); ok {
			maxRows = v
		}
	}

	// Start from current position
	startRow := rs.CurrentRow
	if startRow < 0 {
		startRow = 0
	}

	endRow := len(rs.allData)
	if maxRows > 0 && startRow+maxRows < endRow {
		endRow = startRow + maxRows
	}

	// Build 2D array [field][record]
	// In classic ADO, GetRows returns array(field, record)
	if len(rs.columns) == 0 && len(rs.allData) > 0 {
		// Extract column names from first row
		for key := range rs.allData[0] {
			rs.columns = append(rs.columns, key)
		}
	}

	numFields := len(rs.columns)
	numRecords := endRow - startRow

	// Create 2D array: [field][record]
	result := make([][]interface{}, numFields)
	for i := range result {
		result[i] = make([]interface{}, numRecords)
	}

	// Fill the array
	for recIdx := 0; recIdx < numRecords; recIdx++ {
		row := rs.allData[startRow+recIdx]
		for fieldIdx, colName := range rs.columns {
			result[fieldIdx][recIdx] = row[strings.ToLower(colName)]
		}
	}

	// Move current position to EOF
	rs.CurrentRow = endRow
	if rs.CurrentRow >= len(rs.allData) {
		rs.EOF = true
	}

	return result
}

// applySorting sorts the recordset by the specified field
func (rs *ADODBRecordset) applySorting(sortStr string) {
	if len(rs.allData) == 0 {
		return
	}

	// Parse sort string: "FieldName [ASC|DESC]"
	parts := strings.Fields(strings.TrimSpace(sortStr))
	if len(parts) == 0 {
		return
	}

	rs.SortField = strings.ToLower(parts[0])
	rs.SortOrder = "ASC"
	if len(parts) > 1 {
		rs.SortOrder = strings.ToUpper(parts[1])
	}

	// Sort the data while keeping original snapshots aligned
	type rowPair struct {
		current  map[string]interface{}
		original map[string]interface{}
	}
	pairs := make([]rowPair, len(rs.allData))
	for i := range rs.allData {
		pairs[i] = rowPair{current: rs.allData[i], original: rs.originalData[i]}
	}

	sort.Slice(pairs, func(i, j int) bool {
		valI := pairs[i].current[rs.SortField]
		valJ := pairs[j].current[rs.SortField]

		// Compare based on type
		var less bool
		switch vI := valI.(type) {
		case int:
			if vJ, ok := valJ.(int); ok {
				less = vI < vJ
			}
		case int64:
			if vJ, ok := valJ.(int64); ok {
				less = vI < vJ
			}
		case float64:
			if vJ, ok := valJ.(float64); ok {
				less = vI < vJ
			}
		case string:
			if vJ, ok := valJ.(string); ok {
				less = vI < vJ
			}
		default:
			// Try string comparison
			less = fmt.Sprintf("%v", valI) < fmt.Sprintf("%v", valJ)
		}

		if rs.SortOrder == "DESC" {
			return !less
		}
		return less
	})

	for i := range pairs {
		rs.allData[i] = pairs[i].current
		rs.originalData[i] = pairs[i].original
	}

	// Update current position if valid
	if rs.CurrentRow >= 0 && rs.CurrentRow < len(rs.allData) {
		rs.fetchCurrentRow()
	}
}

// findRecord searches for a record matching the criteria
// Criteria format: "FieldName = 'Value'" or "FieldName = Value"
func (rs *ADODBRecordset) findRecord(criteria string) {
	if len(rs.allData) == 0 {
		rs.EOF = true
		return
	}

	// Simple parser: "Field = Value"
	parts := strings.SplitN(criteria, "=", 2)
	if len(parts) != 2 {
		return
	}

	fieldName := strings.ToLower(strings.TrimSpace(parts[0]))
	searchValue := strings.TrimSpace(parts[1])
	// Remove quotes if present
	searchValue = strings.Trim(searchValue, "'\"")

	// Start searching from current position + 1
	startPos := rs.CurrentRow + 1
	if startPos < 0 {
		startPos = 0
	}

	for i := startPos; i < len(rs.allData); i++ {
		row := rs.allData[i]
		if val, ok := row[fieldName]; ok {
			// Compare values
			valStr := fmt.Sprintf("%v", val)
			if strings.EqualFold(valStr, searchValue) {
				// Found match
				rs.CurrentRow = i
				rs.currentData = row
				rs.updateFieldsCollection()
				rs.EOF = false
				rs.BOF = false
				return
			}
		}
	}

	// Not found, set EOF
	rs.EOF = true
}

// calculatePageCount calculates total pages based on RecordCount and PageSize
func (rs *ADODBRecordset) calculatePageCount() {
	if rs.PageSize <= 0 {
		rs.PageCount = 0
		return
	}
	rs.PageCount = (rs.RecordCount + rs.PageSize - 1) / rs.PageSize
}

// moveToAbsolutePage moves current position to the specified page
func (rs *ADODBRecordset) moveToAbsolutePage() {
	if rs.PageSize <= 0 || rs.AbsolutePage <= 0 {
		return
	}

	// Calculate starting row for the page (1-based page number)
	startRow := (rs.AbsolutePage - 1) * rs.PageSize

	if startRow >= len(rs.allData) {
		// Page out of range, move to last valid row
		rs.CurrentRow = len(rs.allData) - 1
		rs.EOF = true
	} else {
		rs.CurrentRow = startRow
		rs.EOF = false
		rs.BOF = false
		rs.fetchCurrentRow()
	}
}

// applyFilter applies a filter criteria to the recordset
// Creates a temporary view by filtering allData based on the criteria
// Filter format: "FieldName = 'Value'" or "FieldName > Value" or "FieldName LIKE 'pattern'"
func (rs *ADODBRecordset) applyFilter(criteria string) {
	if criteria == "" || len(rs.allData) == 0 {
		rs.clearFilter()
		return
	}

	rs.FilterCriteria = criteria
	rs.filteredIndices = make([]int, 0)

	// Parse simple criteria: "Field operator Value"
	// Support: =, <>, >, <, >=, <=, LIKE
	criteria = strings.TrimSpace(criteria)

	// Simple parser - supports basic comparisons
	for i, row := range rs.allData {
		if rs.matchesFilter(row, criteria) {
			rs.filteredIndices = append(rs.filteredIndices, i)
		}
	}

	rs.isFiltered = true

	// Move to first filtered record
	if len(rs.filteredIndices) > 0 {
		rs.CurrentRow = rs.filteredIndices[0]
		rs.currentData = rs.allData[rs.CurrentRow]
		rs.updateFieldsCollection()
		rs.EOF = false
		rs.BOF = false
	} else {
		rs.EOF = true
		rs.BOF = true
	}
}

// clearFilter removes the filter and restores full recordset view
func (rs *ADODBRecordset) clearFilter() {
	rs.FilterCriteria = ""
	rs.filteredIndices = nil
	rs.isFiltered = false

	// Move to first record
	if len(rs.allData) > 0 {
		rs.CurrentRow = 0
		rs.currentData = rs.allData[0]
		rs.updateFieldsCollection()
		rs.EOF = false
		rs.BOF = false
	}
}

// matchesFilter checks if a row matches the filter criteria
func (rs *ADODBRecordset) matchesFilter(row map[string]interface{}, criteria string) bool {
	// Support basic operators: =, <>, >, <, >=, <=
	operators := []string{">=", "<=", "<>", "=", ">", "<"}

	for _, op := range operators {
		if idx := strings.Index(criteria, op); idx > 0 {
			fieldName := strings.ToLower(strings.TrimSpace(criteria[:idx]))
			valueStr := strings.TrimSpace(criteria[idx+len(op):])
			valueStr = strings.Trim(valueStr, "'\"")

			fieldValue, exists := row[fieldName]
			if !exists {
				return false
			}

			return rs.compareValues(fieldValue, op, valueStr)
		}
	}

	// Support LIKE operator
	if strings.Contains(strings.ToUpper(criteria), " LIKE ") {
		parts := strings.SplitN(strings.ToUpper(criteria), " LIKE ", 2)
		if len(parts) == 2 {
			fieldName := strings.ToLower(strings.TrimSpace(parts[0]))
			pattern := strings.TrimSpace(parts[1])
			pattern = strings.Trim(pattern, "'\"")

			fieldValue, exists := row[fieldName]
			if !exists {
				return false
			}

			// Simple pattern matching: % as wildcard
			pattern = strings.ReplaceAll(pattern, "%", ".*")
			valueStr := fmt.Sprintf("%v", fieldValue)
			matched := strings.Contains(strings.ToLower(valueStr), strings.ToLower(strings.ReplaceAll(pattern, ".*", "")))
			return matched
		}
	}

	return false
}

// compareValues compares two values using the specified operator
func (rs *ADODBRecordset) compareValues(fieldValue interface{}, operator, compareStr string) bool {
	fieldStr := fmt.Sprintf("%v", fieldValue)

	switch operator {
	case "=":
		return strings.EqualFold(fieldStr, compareStr)
	case "<>":
		return !strings.EqualFold(fieldStr, compareStr)
	case ">":
		// Try numeric comparison first
		if fVal, err := fmt.Sscanf(fieldStr, "%f", new(float64)); err == nil && fVal == 1 {
			var fField, fCompare float64
			fmt.Sscanf(fieldStr, "%f", &fField)
			fmt.Sscanf(compareStr, "%f", &fCompare)
			return fField > fCompare
		}
		return fieldStr > compareStr
	case "<":
		if fVal, err := fmt.Sscanf(fieldStr, "%f", new(float64)); err == nil && fVal == 1 {
			var fField, fCompare float64
			fmt.Sscanf(fieldStr, "%f", &fField)
			fmt.Sscanf(compareStr, "%f", &fCompare)
			return fField < fCompare
		}
		return fieldStr < compareStr
	case ">=":
		if fVal, err := fmt.Sscanf(fieldStr, "%f", new(float64)); err == nil && fVal == 1 {
			var fField, fCompare float64
			fmt.Sscanf(fieldStr, "%f", &fField)
			fmt.Sscanf(compareStr, "%f", &fCompare)
			return fField >= fCompare
		}
		return fieldStr >= compareStr
	case "<=":
		if fVal, err := fmt.Sscanf(fieldStr, "%f", new(float64)); err == nil && fVal == 1 {
			var fField, fCompare float64
			fmt.Sscanf(fieldStr, "%f", &fField)
			fmt.Sscanf(compareStr, "%f", &fCompare)
			return fField <= fCompare
		}
		return fieldStr <= compareStr
	}

	return false
}

// findCurrentFilteredIndex returns the index in filteredIndices array for current row
func (rs *ADODBRecordset) findCurrentFilteredIndex() int {
	for i, idx := range rs.filteredIndices {
		if idx == rs.CurrentRow {
			return i
		}
	}
	return -1
}

func (rs *ADODBRecordset) setCurrentRowByIndex(index int) {
	if rs.isFiltered {
		if len(rs.filteredIndices) == 0 {
			rs.EOF = true
			rs.BOF = true
			return
		}
		if index < 0 {
			rs.BOF = true
			rs.EOF = false
			return
		}
		if index >= len(rs.filteredIndices) {
			rs.EOF = true
			rs.BOF = false
			return
		}
		rs.CurrentRow = rs.filteredIndices[index]
		rs.fetchCurrentRow()
		rs.BOF = false
		rs.EOF = false
		return
	}

	if len(rs.allData) == 0 {
		rs.EOF = true
		rs.BOF = true
		return
	}
	if index < 0 {
		rs.BOF = true
		rs.EOF = false
		return
	}
	if index >= len(rs.allData) {
		rs.EOF = true
		rs.BOF = false
		return
	}
	if rs.CurrentRow != index {
		rs.CurrentRow = index
	}
	rs.fetchCurrentRow()
	rs.BOF = false
	rs.EOF = false
}

// supportsFeature checks if the recordset supports a specific cursor feature
// Common constants from ADO:
// adAddNew = 0x1000400
// adApproxPosition = 0x4000
// adBookmark = 0x2000
// adDelete = 0x1000800
// adFind = 0x80000
// adHoldRecords = 0x100
// adMovePrevious = 0x200
// adResync = 0x20000
// adUpdate = 0x1008000
// adUpdateBatch = 0x10000
func (rs *ADODBRecordset) supportsFeature(option int) bool {
	// For this generic SQL implementation, we support most common features
	switch option {
	case 0x1000400: // adAddNew
		return true
	case 0x1000800: // adDelete
		return true
	case 0x1008000: // adUpdate
		return true
	case 0x200: // adMovePrevious
		return true
	case 0x80000: // adFind
		return true
	case 0x2000: // adBookmark
		return false // Not implemented yet
	case 0x4000: // adApproxPosition
		return true
	case 0x10000: // adUpdateBatch
		return false // Not implemented yet
	case 0x100: // adHoldRecords
		return true
	case 0x20000: // adResync
		return false // Not implemented yet
	default:
		// By default, return true for common operations
		return true
	}
}

// performSQLUpdate attempts to persist changes to the database
func (rs *ADODBRecordset) performSQLUpdate() {
	tableName := extractTableName(rs.Source)
	if tableName == "" {
		return
	}

	driver := ""
	if rs.dbConn != nil {
		driver = rs.dbConn.dbDriver
	}

	if rs.newData != nil {
		// INSERT
		var cols []string
		var placeholders []string
		var values []interface{}

		for col, val := range rs.newData {
			cols = append(cols, col)
			placeholders = append(placeholders, "?")
			values = append(values, val)
		}

		if len(cols) == 0 {
			return
		}

		query := fmt.Sprintf("INSERT INTO %s (%s) VALUES (%s)",
			tableName,
			strings.Join(cols, ", "),
			strings.Join(placeholders, ", "))
		query = rewritePlaceholders(query, driver)

		var result sql.Result
		var err error
		if rs.dbConn != nil {
			result, err = rs.dbConn.execStatement(query, values)
		} else {
			result, err = rs.db.Exec(query, values...)
		}

		if err != nil {
			if rs.ctx != nil && rs.ctx.Err != nil {
				rs.ctx.Err.SetError(err)
			}
		} else {
			// Try to retrieve LastInsertId and update identity field
			if lastId, err := result.LastInsertId(); err == nil && lastId > 0 {
				// Look for standard ID field names
				idFields := []string{"id", "iid", "identity", "autoid"}

				// Also try to find first integer field if specific ID not found? No, unsafe.

				for _, idField := range idFields {
					found := false
					if rs.newData != nil {
						if _, ok := rs.newData[idField]; ok || strings.EqualFold(idField, "id") {
							rs.newData[idField] = lastId
							found = true
						}
					}
					for _, f := range rs.Fields.fields {
						if strings.EqualFold(f.Name, idField) {
							f.Value = lastId
							if rs.currentData != nil {
								rs.currentData[strings.ToLower(f.Name)] = lastId
							}
							if rs.CurrentRow >= 0 && rs.CurrentRow < len(rs.allData) {
								rs.allData[rs.CurrentRow][strings.ToLower(f.Name)] = lastId
							}
							found = true
							break
						}
					}
					if found {
						break
					}
				}
			}
		}
	} else {
		// UPDATE
		if rs.currentData == nil {
			return
		}
		whereClause := cleanWhereClause(extractWhereClause(rs.Source))
		whereParams := []interface{}{}
		if whereClause == "" {
			if rs.CurrentRow < 0 || rs.CurrentRow >= len(rs.originalData) {
				return
			}
			whereClause, whereParams = buildRowWhereClause(rs.originalData[rs.CurrentRow])
			if whereClause == "" {
				return
			}
		}

		var sets []string
		var values []interface{}

		for col, val := range rs.currentData {
			sets = append(sets, fmt.Sprintf("%s = ?", col))
			values = append(values, val)
		}

		if len(sets) == 0 {
			return
		}

		values = append(values, whereParams...)

		query := fmt.Sprintf("UPDATE %s SET %s %s", tableName, strings.Join(sets, ", "), whereClause)
		query = rewritePlaceholders(query, driver)

		var err error
		if rs.dbConn != nil {
			_, err = rs.dbConn.execStatement(query, values)
		} else {
			_, err = rs.db.Exec(query, values...)
		}

		if err != nil {
			if rs.ctx != nil && rs.ctx.Err != nil {
				rs.ctx.Err.SetError(err)
			}
			return
		}

		if rs.CurrentRow >= 0 && rs.CurrentRow < len(rs.originalData) {
			rs.originalData[rs.CurrentRow] = cloneRow(rs.currentData)
		}
	}
}

func (rs *ADODBRecordset) performSQLDelete() {
	tableName := extractTableName(rs.Source)
	if tableName == "" {
		return
	}

	driver := ""
	if rs.dbConn != nil {
		driver = rs.dbConn.dbDriver
	}

	whereClause := cleanWhereClause(extractWhereClause(rs.Source))
	whereParams := []interface{}{}
	if whereClause == "" {
		if rs.CurrentRow < 0 || rs.CurrentRow >= len(rs.originalData) {
			return
		}
		whereClause, whereParams = buildRowWhereClause(rs.originalData[rs.CurrentRow])
		if whereClause == "" {
			return
		}
	}

	query := fmt.Sprintf("DELETE FROM %s %s", tableName, whereClause)
	query = rewritePlaceholders(query, driver)

	var err error
	if rs.dbConn != nil {
		_, err = rs.dbConn.execStatement(query, whereParams)
	} else {
		_, err = rs.db.Exec(query, whereParams...)
	}
	if err != nil {
		if rs.ctx != nil && rs.ctx.Err != nil {
			rs.ctx.Err.SetError(err)
		}
	}
}

func extractTableName(sql string) string {
	upperSQL := strings.ToUpper(sql)
	idx := strings.Index(upperSQL, " FROM ")
	if idx == -1 {
		return ""
	}
	remaining := strings.TrimSpace(sql[idx+6:])
	parts := strings.Fields(remaining)
	if len(parts) > 0 {
		return parts[0]
	}
	return ""
}

func extractWhereClause(sql string) string {
	upperSQL := strings.ToUpper(sql)
	idx := strings.Index(upperSQL, " WHERE ")
	if idx == -1 {
		return ""
	}
	return sql[idx:]
}

func cleanWhereClause(whereClause string) string {
	upper := strings.ToUpper(whereClause)
	for _, token := range []string{" ORDER BY ", " LIMIT ", " OFFSET ", " FETCH ", " FOR "} {
		if idx := strings.Index(upper, token); idx != -1 {
			return strings.TrimSpace(whereClause[:idx])
		}
	}
	return strings.TrimSpace(whereClause)
}

func buildRowWhereClause(row map[string]interface{}) (string, []interface{}) {
	if row == nil {
		return "", nil
	}
	if idValue, ok := row["id"]; ok {
		if idValue == nil {
			return "WHERE id IS NULL", nil
		}
		return "WHERE id = ?", []interface{}{idValue}
	}

	clauses := make([]string, 0, len(row))
	params := make([]interface{}, 0, len(row))
	for col, val := range row {
		if val == nil {
			clauses = append(clauses, fmt.Sprintf("%s IS NULL", col))
			continue
		}
		clauses = append(clauses, fmt.Sprintf("%s = ?", col))
		params = append(params, val)
	}
	if len(clauses) == 0 {
		return "", nil
	}
	return "WHERE " + strings.Join(clauses, " AND "), params
}

// --- ADODB.Recordset (OLE Wrapper) ---

// ADODBOLERecordset wraps an OLE/COM Recordset object for Access databases
type ADODBOLERecordset struct {
	oleRecordset     *ole.IDispatch
	activeConnection *ole.IDispatch // Store the ActiveConnection
	ctx              *ExecutionContext
}

// NewADODBOLERecordset creates a new OLE recordset wrapper
func NewADODBOLERecordset(oleRs *ole.IDispatch, ctx *ExecutionContext) *ADODBOLERecordset {
	return &ADODBOLERecordset{
		oleRecordset: oleRs,
		ctx:          ctx,
	}
}

func (rs *ADODBOLERecordset) getFieldValue(fieldName interface{}) (interface{}, bool) {
	if rs.oleRecordset == nil {
		return nil, false
	}
	fieldsResult, err := oleutil.GetProperty(rs.oleRecordset, "Fields")
	if err != nil {
		return nil, false
	}
	fieldsDisp := fieldsResult.ToIDispatch()
	if fieldsDisp == nil {
		return nil, false
	}
	defer fieldsDisp.Release()

	fieldResult, err := oleutil.GetProperty(fieldsDisp, "Item", fieldName)
	if err != nil {
		return nil, false
	}
	fieldDisp := fieldResult.ToIDispatch()
	if fieldDisp == nil {
		return nil, false
	}
	defer fieldDisp.Release()

	valueResult, err := oleutil.GetProperty(fieldDisp, "Value")
	if err != nil {
		return nil, false
	}
	return valueResult.Value(), true
}

func (rs *ADODBOLERecordset) setFieldValue(fieldName interface{}, value interface{}) error {
	if rs.oleRecordset == nil {
		return fmt.Errorf("Recordset is closed or invalid")
	}

	// Try to convert numeric values to int32 if they are whole numbers and fit, to help Access with Integer fields (VT_I4)
	switch v := value.(type) {
	case float64:
		if v == float64(int32(v)) {
			value = int32(v)
		}
	case int:
		value = int32(v)
	case int64:
		value = int32(v)
	}

	fieldsResult, err := oleutil.GetProperty(rs.oleRecordset, "Fields")
	if err != nil {
		return err
	}
	fieldsDisp := fieldsResult.ToIDispatch()
	if fieldsDisp == nil {
		return fmt.Errorf("Fields collection is null")
	}
	defer fieldsDisp.Release()

	fieldResult, err := oleutil.GetProperty(fieldsDisp, "Item", fieldName)
	if err != nil {
		return err
	}
	fieldDisp := fieldResult.ToIDispatch()
	if fieldDisp == nil {
		return fmt.Errorf("Field '%v' not found", fieldName)
	}
	defer fieldDisp.Release()

	value = coerceOLEFieldValue(fieldDisp, value)

	if _, err := oleutil.PutProperty(fieldDisp, "Value", value); err != nil {
		fmt.Printf("[ADODB ERROR] setFieldValue failed for field '%v', value '%v' (Type: %T). Error: %v\n", fieldName, value, value, err)
		return err
	}
	return nil
}

func coerceOLEFieldValue(fieldDisp *ole.IDispatch, value interface{}) interface{} {
	if fieldDisp == nil {
		return value
	}

	typeResult, err := oleutil.GetProperty(fieldDisp, "Type")
	if err != nil {
		return value
	}
	fieldType := toInt(typeResult.Value())

	switch fieldType {
	case 7, 133, 134, 135:
		return coerceOLEDateValue(value)
	case 2, 3, 16, 17, 18, 19, 20, 21:
		return toInt32Variant(value)
	case 4, 5, 6, 14, 131:
		return toFloat(value)
	default:
		return value
	}
}

func coerceOLEDateValue(value interface{}) interface{} {
	switch v := value.(type) {
	case time.Time:
		return v
	case string:
		if parsed, ok := parseOLEDateString(v); ok {
			return parsed
		}
	}
	return value
}

func parseOLEDateString(input string) (time.Time, bool) {
	value := strings.TrimSpace(input)
	if value == "" {
		return time.Time{}, false
	}

	layouts := []string{
		"02/01/2006",
		"02/01/2006 15:04:05",
		"02/01/2006 15:04",
		"2006-01-02",
		"2006-01-02 15:04:05",
		"2006-01-02 15:04",
		"01/02/2006",
		"01/02/2006 15:04:05",
		"01/02/2006 15:04",
	}

	for _, layout := range layouts {
		if parsed, err := time.ParseInLocation(layout, value, time.Local); err == nil {
			return parsed, true
		}
	}

	return time.Time{}, false
}

func (rs *ADODBOLERecordset) GetProperty(name string) interface{} {
	if rs.oleRecordset == nil {
		return nil
	}

	defer func() {
		if r := recover(); r != nil {
			// Silently recover from OLE errors
		}
	}()

	switch strings.ToLower(name) {
	case "eof":
		result, err := oleutil.GetProperty(rs.oleRecordset, "EOF")
		if err == nil {
			return result.Value()
		}
	case "bof":
		result, err := oleutil.GetProperty(rs.oleRecordset, "BOF")
		if err == nil {
			return result.Value()
		}
	case "recordcount":
		result, err := oleutil.GetProperty(rs.oleRecordset, "RecordCount")
		if err == nil {
			return result.Value()
		}
	case "fields":
		result, err := oleutil.GetProperty(rs.oleRecordset, "Fields")
		if err == nil {
			fieldsDisp := result.ToIDispatch()
			if fieldsDisp != nil {
				// Return the ASPLibrary wrapper, not the raw OLE object
				return NewADOOLEFields(NewADODBOLEFields(fieldsDisp, rs))
			} else {
				return nil
			}
		}
	case "absoluteposition":
		result, err := oleutil.GetProperty(rs.oleRecordset, "AbsolutePosition")
		if err == nil {
			return result.Value()
		}
	case "pagesize":
		result, err := oleutil.GetProperty(rs.oleRecordset, "PageSize")
		if err == nil {
			return result.Value()
		}
	case "absolutepage":
		result, err := oleutil.GetProperty(rs.oleRecordset, "AbsolutePage")
		if err == nil {
			return result.Value()
		}
	case "state":
		result, err := oleutil.GetProperty(rs.oleRecordset, "State")
		if err == nil {
			return result.Value()
		}
	case "cursorlocation":
		result, err := oleutil.GetProperty(rs.oleRecordset, "CursorLocation")
		if err == nil {
			return result.Value()
		}
	}
	return nil
}

func (rs *ADODBOLERecordset) SetProperty(name string, value interface{}) {
	if rs.oleRecordset == nil {
		return
	}

	switch strings.ToLower(name) {
	case "activeconnection":
		// Store the connection - it might be an ASPLibrary wrapper or OLE IDispatch
		if conn, ok := value.(*ADODBConnection); ok {
			// It's our wrapper, extract the OLE connection
			rs.activeConnection = conn.oleConnection
		} else if conn, ok := value.(*ADOConnection); ok {
			// It's the ASPLibrary wrapper
			rs.activeConnection = conn.lib.oleConnection
		} else if conn, ok := value.(*ole.IDispatch); ok {
			// It's already an OLE IDispatch
			rs.activeConnection = conn
		}
	case "pagesize":
		oleutil.PutProperty(rs.oleRecordset, "PageSize", toInt32Variant(value))
	case "absolutepage":
		oleutil.PutProperty(rs.oleRecordset, "AbsolutePage", toInt32Variant(value))
	case "absoluteposition":
		oleutil.PutProperty(rs.oleRecordset, "AbsolutePosition", toInt32Variant(value))
	case "cursortype":
		oleutil.PutProperty(rs.oleRecordset, "CursorType", toInt32Variant(value))
	case "locktype":
		oleutil.PutProperty(rs.oleRecordset, "LockType", toInt32Variant(value))
	case "cursorlocation":
		oleutil.PutProperty(rs.oleRecordset, "CursorLocation", toInt32Variant(value))
	default:
		if err := rs.setFieldValue(name, value); err != nil {
			panic(fmt.Errorf("ADODB.Recordset: Failed to set field '%s'. Details: %v", name, err))
		}
	}
}

func (rs *ADODBOLERecordset) CallMethod(name string, args ...interface{}) interface{} {
	if rs.oleRecordset == nil {
		return nil

	}

	method := strings.ToLower(name)
	if method == "" && len(args) > 0 {
		value, _ := rs.getFieldValue(args[0])
		return value
	}

	switch method {
	case "fields":
		// Handle rs.Fields("fieldName") shortcut - common VBScript pattern
		if len(args) > 0 {
			// Get the Fields collection
			result, err := oleutil.GetProperty(rs.oleRecordset, "Fields")
			if err != nil {
				return nil
			}
			fieldsDisp := result.ToIDispatch()
			if fieldsDisp != nil {
				defer fieldsDisp.Release()
				// Get the field by name/index
				fieldResult, err := oleutil.GetProperty(fieldsDisp, "Item", args[0])
				if err != nil {
					fmt.Printf("Error> ADODB.Recordset CallMethod Fields Item error: %s\n", err.Error())
					return nil
				}
				fieldDisp := fieldResult.ToIDispatch()
				if fieldDisp != nil {
					defer fieldDisp.Release()
					nameResult, nameErr := oleutil.GetProperty(fieldDisp, "Name")
					if nameErr != nil {
						fmt.Printf("Error> ADODB.Recordset CallMethod Fields Item Name error: %s\n", nameErr)
						return nil
					}
					fieldName := nameResult.ToString()
					return &OLEFieldProxy{recordset: rs, name: fieldName}
				}
			}
		}
		return nil
	case "movenext":
		oleutil.CallMethod(rs.oleRecordset, "MoveNext")
		return nil
	case "moveprevious":
		oleutil.CallMethod(rs.oleRecordset, "MovePrevious")
		return nil
	case "movefirst":
		oleutil.CallMethod(rs.oleRecordset, "MoveFirst")
		return nil
	case "movelast":
		oleutil.CallMethod(rs.oleRecordset, "MoveLast")
		return nil
	case "move":
		if len(args) > 0 {
			oleutil.CallMethod(rs.oleRecordset, "Move", args[0])
		}
		return nil
	case "open":
		if len(args) > 0 {
			// If ActiveConnection was set and only SQL is provided, add the connection as 2nd arg
			// Force cursorType=3 (static) and lockType=3 (optimistic) for Access compatibility
			if rs.activeConnection != nil && len(args) == 1 {
				// args[0] = source (SQL)
				oleutil.CallMethod(rs.oleRecordset, "Open", args[0], rs.activeConnection, int32(3), int32(3))
			} else if rs.activeConnection != nil && len(args) == 2 {
				// If activeConnection was already passed, use it; add cursor/lock types
				oleutil.CallMethod(rs.oleRecordset, "Open", args[0], rs.activeConnection, int32(3), int32(3))
			} else {
				// Use whatever was passed, but coerce cursor/lock numeric arguments to int32
				openArgs := make([]interface{}, len(args))
				copy(openArgs, args)
				if len(openArgs) >= 3 {
					openArgs[2] = toInt32Variant(openArgs[2])
				}
				if len(openArgs) >= 4 {
					openArgs[3] = toInt32Variant(openArgs[3])
				}
				oleutil.CallMethod(rs.oleRecordset, "Open", openArgs...)
			}
		}
		return nil
	case "close":
		oleutil.CallMethod(rs.oleRecordset, "Close")
		if rs.oleRecordset != nil {
			rs.oleRecordset.Release()
			rs.oleRecordset = nil
		}
		return nil
	case "addnew":
		if _, err := oleutil.CallMethod(rs.oleRecordset, "AddNew"); err != nil {
			panic(fmt.Errorf("ADODB.Recordset: AddNew failed. %v", err))
		}
		return nil
	case "update":
		var err error
		if len(args) > 0 {
			_, err = oleutil.CallMethod(rs.oleRecordset, "Update", args...)
		} else {
			_, err = oleutil.CallMethod(rs.oleRecordset, "Update")
		}
		if err != nil {
			fmt.Printf("[ADODB ERROR] Update failed. Error: %v\n", err)
			panic(fmt.Errorf("ADODB.Recordset: Update failed. %v", err))
		}
		return nil
	case "delete":
		if _, err := oleutil.CallMethod(rs.oleRecordset, "Delete"); err != nil {
			panic(fmt.Errorf("ADODB.Recordset: Delete failed. %v", err))
		}
		return nil
	case "cancelupdate":
		oleutil.CallMethod(rs.oleRecordset, "CancelUpdate")
		return nil
	case "getrows":
		// GetRows returns a 2D array [field][record] from current position
		return rs.getRows(args)
	}

	return nil
}

// getRows returns a 2D array containing all records from current position to end
// Returns array(field, record) - same as classic ADO
func (rs *ADODBOLERecordset) getRows(args []interface{}) interface{} {
	if rs.oleRecordset == nil {
		return [][]interface{}{}
	}

	// Check if EOF
	eofResult, err := oleutil.GetProperty(rs.oleRecordset, "EOF")
	if err != nil {
		return [][]interface{}{}
	}
	if eofResult.Value() == true {
		return [][]interface{}{}
	}

	// Get Fields collection to know column count and names
	fieldsResult, err := oleutil.GetProperty(rs.oleRecordset, "Fields")
	if err != nil {
		return [][]interface{}{}
	}
	fieldsDisp := fieldsResult.ToIDispatch()
	if fieldsDisp == nil {
		return [][]interface{}{}
	}
	defer fieldsDisp.Release()

	countResult, err := oleutil.GetProperty(fieldsDisp, "Count")
	if err != nil {
		return [][]interface{}{}
	}
	fieldCount := int(countResult.Value().(int32))
	if fieldCount == 0 {
		return [][]interface{}{}
	}

	// Determine how many rows to fetch
	maxRows := -1 // -1 means all remaining rows
	if len(args) > 0 {
		switch v := args[0].(type) {
		case int:
			maxRows = v
		case int32:
			maxRows = int(v)
		case int64:
			maxRows = int(v)
		case float64:
			maxRows = int(v)
		}
	}

	// Collect all rows from current position
	var allRows [][]interface{}
	rowCount := 0
	for {
		// Check EOF
		eofCheck, _ := oleutil.GetProperty(rs.oleRecordset, "EOF")
		if eofCheck.Value() == true {
			break
		}

		// Check row limit
		if maxRows > 0 && rowCount >= maxRows {
			break
		}

		// Read current row
		rowData := make([]interface{}, fieldCount)
		for i := 0; i < fieldCount; i++ {
			fieldResult, err := oleutil.GetProperty(fieldsDisp, "Item", i)
			if err != nil {
				rowData[i] = nil
				continue
			}
			fieldDisp := fieldResult.ToIDispatch()
			if fieldDisp != nil {
				valueResult, err := oleutil.GetProperty(fieldDisp, "Value")
				if err == nil {
					rowData[i] = valueResult.Value()
				} else {
					rowData[i] = nil
				}
				fieldDisp.Release()
			}
		}
		allRows = append(allRows, rowData)
		rowCount++

		// Move to next record
		oleutil.CallMethod(rs.oleRecordset, "MoveNext")
	}

	if len(allRows) == 0 {
		return [][]interface{}{}
	}

	// Transpose to [field][record] format (classic ADO GetRows format)
	numRecords := len(allRows)
	result := make([][]interface{}, fieldCount)
	for i := range result {
		result[i] = make([]interface{}, numRecords)
	}

	for recIdx, row := range allRows {
		for fieldIdx := 0; fieldIdx < fieldCount; fieldIdx++ {
			if fieldIdx < len(row) {
				result[fieldIdx][recIdx] = row[fieldIdx]
			}
		}
	}

	return result
}

// --- ADODB.Fields (OLE Wrapper) ---

// ADODBOLEFields wraps an OLE/COM Fields collection
type ADODBOLEFields struct {
	oleFields *ole.IDispatch
	parent    *ADODBOLERecordset
}

// NewADODBOLEFields creates a new OLE fields wrapper
func NewADODBOLEFields(oleFields *ole.IDispatch, parent *ADODBOLERecordset) *ADODBOLEFields {
	return &ADODBOLEFields{
		oleFields: oleFields,
		parent:    parent,
	}
}

func (f *ADODBOLEFields) GetProperty(name string) interface{} {
	if f.oleFields == nil {
		return nil
	}

	defer func() {
		if r := recover(); r != nil {
			// Silently recover from OLE errors
			fmt.Printf("[DEBUG]: OLE recovered from panic in GetProperty: %v\n", r)
		}
	}()

	nameLower := strings.ToLower(name)
	switch nameLower {
	case "count":
		result, err := oleutil.GetProperty(f.oleFields, "Count")
		if err == nil {
			return result.Value()
		}
	case "item":
		// "Item" is not a field property - it's a method
		// Return nil to indicate this should be handled as CallMethod instead
		return nil
	default:
		// Try to access as a field name directly (subscript access)
		result, err := oleutil.GetProperty(f.oleFields, "Item", name)
		if err != nil {
			return nil
		}
		fieldDisp := result.ToIDispatch()
		if fieldDisp != nil {
			defer fieldDisp.Release()
			nameResult, nameErr := oleutil.GetProperty(fieldDisp, "Name")
			if nameErr != nil {
				return nil
			}
			fieldName := nameResult.ToString()
			return &OLEFieldProxy{recordset: f.parent, name: fieldName}
		}
	}
	return nil
}

func (f *ADODBOLEFields) SetProperty(name string, value interface{}) {}

func (f *ADODBOLEFields) CallMethod(name string, args ...interface{}) interface{} {
	if f.oleFields == nil {
		return nil
	}

	defer func() {
		if r := recover(); r != nil {
			// Silently recover from OLE errors
		}
	}()

	nameLower := strings.ToLower(name)
	// Default method is "Item" - empty name means default dispatch
	if nameLower == "" || nameLower == "item" {
		if len(args) > 0 {
			result, err := oleutil.GetProperty(f.oleFields, "Item", args[0])
			if err != nil {
				return nil
			}
			fieldDisp := result.ToIDispatch()
			if fieldDisp != nil {
				defer fieldDisp.Release()
				nameResult, nameErr := oleutil.GetProperty(fieldDisp, "Name")
				if nameErr != nil {
					return nil
				}
				fieldName := nameResult.ToString()
				return &OLEFieldProxy{recordset: f.parent, name: fieldName}
			}
		}
		return nil
	}

	switch nameLower {
	case "count":
		result, err := oleutil.GetProperty(f.oleFields, "Count")
		if err == nil {
			return result.Value()
		}
	}
	return nil
}

// Enumeration returns all Field proxies for For Each support.
func (f *ADODBOLEFields) Enumeration() []interface{} {
	if f.oleFields == nil {
		return []interface{}{}
	}

	countResult, err := oleutil.GetProperty(f.oleFields, "Count")
	if err != nil {
		return []interface{}{}
	}
	count := toInt(countResult.Value())
	if count < 0 {
		count = 0
	}

	items := make([]interface{}, 0, count)
	for i := 0; i < count; i++ {
		itemResult, err := oleutil.GetProperty(f.oleFields, "Item", i)
		if err != nil {
			continue
		}
		fieldDisp := itemResult.ToIDispatch()
		if fieldDisp == nil {
			continue
		}
		nameResult, nameErr := oleutil.GetProperty(fieldDisp, "Name")
		fieldDisp.Release()
		if nameErr != nil {
			continue
		}
		fieldName := nameResult.ToString()
		items = append(items, &OLEFieldProxy{recordset: f.parent, name: fieldName})
	}

	return items
}
