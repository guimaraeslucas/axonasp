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
		idx, ok := args[0].(int)
		if !ok || idx < 0 || idx >= len(ec.errors) {
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
}

// NewADODBConnection creates a new connection object
func NewADODBConnection(ctx *ExecutionContext) *ADODBConnection {
	return &ADODBConnection{
		State:  0,
		Mode:   3,
		ctx:    ctx,
		Errors: NewErrorsCollection(),
	}
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
		if v, ok := value.(int); ok {
			c.Mode = v
		}
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
		return c.openDatabase()

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

		sql := fmt.Sprintf("%v", args[0])

		// Handle SQL driver connections
		if c.db != nil {
			result, err := c.db.Exec(sql)
			if err != nil {
				c.Errors.AddError(-1, err.Error(), "ADODB.Connection", "")
				return nil
			}
			affected, _ := result.RowsAffected()
			return int(affected)
		}

		// Handle OLE/Access connections - return Recordset
		if c.oleConnection != nil {
			result, err := oleutil.CallMethod(c.oleConnection, "Execute", sql)
			if err != nil {
				c.Errors.AddError(-1, err.Error(), "ADODB.Connection", "")
				return nil
			}

			// Wrap the OLE Recordset in our ADODBOLERecordset wrapper, then in ASPLibrary wrapper
			if result != nil {
				oleRs := result.ToIDispatch()
				if oleRs != nil {
					// Return the ASPLibrary wrapper, not the raw OLE recordset
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
		c.db.Exec("BEGIN TRANSACTION")
		return nil

	case "committrans":
		if c.db == nil {
			return nil
		}
		c.db.Exec("COMMIT")
		return nil

	case "rollbacktrans":
		if c.db == nil {
			return nil
		}
		c.db.Exec("ROLLBACK")
		return nil
	}

	return nil
}

// openDatabase parses connection string and opens database
func (c *ADODBConnection) openDatabase() interface{} {
	connStr := strings.TrimSpace(c.ConnectionString)
	if connStr == "" {
		return nil
	}

	// Check for Microsoft Access formats - use OLE method
	connStrLower := strings.ToLower(connStr)
	if (strings.Contains(connStrLower, "microsoft.jet.oledb") || strings.Contains(connStrLower, "microsoft.ace.oledb")) && runtime.GOOS == "windows" {
		return c.openAccessDatabase(connStr)
	}

	// Parse ODBC-style connection string for other drivers
	driver, dsn := parseConnectionString(connStr)

	if driver == "" || dsn == "" {
		return nil
	}

	db, err := sql.Open(driver, dsn)
	if err != nil {
		return nil
	}

	// Test connection
	if err := db.Ping(); err != nil {
		db.Close()
		return nil
	}

	c.db = db
	c.dbDriver = driver
	c.State = 1
	return nil
}

// openAccessDatabase opens an Access database using OLE/OLEDB
func (c *ADODBConnection) openAccessDatabase(connStr string) interface{} {
	// Only supported on Windows
	if runtime.GOOS != "windows" {
		fmt.Println("Warning: Direct Access database support is only available on Windows. Please use a different database system for cross-platform compatibility.")
		return nil
	}

	// Try to create ADODB.Connection via OLE
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("ADODB.Connection")
	if err != nil {
		fmt.Printf("Warning: ADODB.Connection COM object cannot be created. Error: %v. Make sure you have Windows COM support and OLEDB drivers installed.\n", err)
		return nil
	}
	defer unknown.Release()

	connection, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		fmt.Println("Warning: Cannot query ADODB.Connection interface. COM support may not be available.")
		return nil
	}

	// Open the connection
	_, err = oleutil.CallMethod(connection, "Open", connStr)
	if err != nil {
		fmt.Printf("Warning: Cannot open Access database. Error details: %v\n", err)
		fmt.Printf("Connection string: %s\n", connStr)
		connection.Release()
		return nil
	}

	// Access database opened successfully

	// Store the OLE connection object and do NOT defer release here
	// It needs to stay alive for the lifetime of the ADODB.Connection object

	c.oleConnection = connection
	c.State = 1
	return nil
}

// parseConnectionString converts ODBC connection string to Go SQL driver and DSN
func parseConnectionString(connStr string) (driver string, dsn string) {
	connStrLower := strings.ToLower(connStr)

	// Handle SQLite formats
	if strings.HasPrefix(connStrLower, "sqlite:") {
		driver = "sqlite"
		dbPath := strings.TrimPrefix(connStrLower, "sqlite:")
		dsn = dbPath
		if dsn == "" {
			dsn = ":memory:"
		}
		return
	}

	// Access databases should be handled via OLE (see openAccessDatabase)
	// If we reach here, Access is not supported on this platform
	if strings.Contains(connStrLower, "microsoft.jet.oledb") || strings.Contains(connStrLower, "microsoft.ace.oledb") {
		fmt.Println("Warning: Direct Access database support is only available on Windows. Please use a different database system for cross-platform compatibility.")
		driver = "sqlite"
		dsn = ":memory:"
		return
	}

	// Parse ODBC-style: Driver={...};Server=...;Database=...;UID=...;PWD=...
	params := make(map[string]string)
	parts := strings.Split(connStrLower, ";")
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if idx := strings.Index(part, "="); idx > 0 {
			key := strings.TrimSpace(part[:idx])
			val := strings.TrimSpace(part[idx+1:])
			// Remove curly braces if present
			val = strings.Trim(val, "{}")
			params[strings.ToLower(key)] = val
		}
	}

	// Detect driver type
	driverStr := params["driver"]
	server := params["server"]
	database := params["database"]
	uid := params["uid"]
	pwd := params["pwd"]

	// MySQL
	if strings.Contains(driverStr, "mysql") {
		driver = "mysql"
		// DSN format: user:password@tcp(server:port)/database
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
		// DSN format: user=user password=pass host=server port=5432 dbname=database sslmode=disable
		port := params["port"]
		if port == "" {
			port = "5432"
		}
		dsn = fmt.Sprintf("user=%s password=%s host=%s port=%s dbname=%s sslmode=disable", uid, pwd, server, port, database)
		return
	}

	// MS SQL Server
	if strings.Contains(driverStr, "sql server") || strings.Contains(driverStr, "mssql") {
		driver = "mssql"
		// DSN format: server=localhost;user id=sa;password=;database=testdb
		port := params["port"]
		if port == "" {
			port = "1433"
		}
		dsn = fmt.Sprintf("server=%s;user id=%s;password=%s;database=%s;port=%s", server, uid, pwd, database, port)
		return
	}

	// Default fallback
	driver = "sqlite"
	dsn = ":memory:"
	return
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
	EOF              bool
	BOF              bool
	RecordCount      int
	State            int // 0 = closed, 1 = open
	CurrentRow       int
	Fields           *FieldsCollection
	rows             *sql.Rows
	db               *sql.DB
	columns          []string
	currentData      map[string]interface{}
	allData          []map[string]interface{}
	newData          map[string]interface{} // For AddNew
	ctx              *ExecutionContext
	PageSize         int    // Number of records per page
	AbsolutePage     int    // Current page number (1-based)
	PageCount        int    // Total number of pages
	SortField        string // Field name for sorting
	SortOrder        string // ASC or DESC
	FilterCriteria   string // Filter expression
	filteredIndices  []int  // Indices of filtered records
	isFiltered       bool   // Whether filter is active
	ActiveConnection interface{} // Stored connection for later use
	CursorType       int    // Cursor type
	LockType         int    // Lock type
}

// NewADODBRecordset creates a new recordset
func NewADODBRecordset(ctx *ExecutionContext) *ADODBRecordset {
	return &ADODBRecordset{
		EOF:          true,
		BOF:          true,
		State:        0,
		CurrentRow:   -1,
		Fields:       NewFieldsCollection(),
		allData:      make([]map[string]interface{}, 0),
		ctx:          ctx,
		PageSize:     10,
		AbsolutePage: 1,
		PageCount:    0,
		SortOrder:    "ASC",
	}
}

func (rs *ADODBRecordset) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
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
	case "fields":
		return rs.Fields
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
	}
	return nil
}

func (rs *ADODBRecordset) SetProperty(name string, value interface{}) {
	switch strings.ToLower(name) {
	case "activeconnection":
		rs.ActiveConnection = value
	case "cursortype":
		if v, ok := value.(int); ok {
			rs.CursorType = v
		}
	case "locktype":
		if v, ok := value.(int); ok {
			rs.LockType = v
		}
	case "pagesize":
		if v, ok := value.(int); ok && v > 0 {
			rs.PageSize = v
			rs.calculatePageCount()
		}
	case "absolutepage":
		if v, ok := value.(int); ok && v > 0 {
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
	}
}

func (rs *ADODBRecordset) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

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
			fmt.Println("Warning: ADODBRecordset.Open called without connection")
			return nil
		}

		return rs.openRecordset(sql, conn)

	case "close":
		if rs.rows != nil {
			rs.rows.Close()
		}
		rs.State = 0
		rs.EOF = true
		rs.BOF = true
		return nil

	case "movenext":
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

	case "addnew":
		rs.newData = make(map[string]interface{})
		return nil

	case "update":
		// Add new row to allData
		if rs.newData != nil {
			rs.allData = append(rs.allData, rs.newData)
			rs.RecordCount = len(rs.allData)
			rs.newData = nil
		}
		return nil

	case "delete":
		if rs.CurrentRow >= 0 && rs.CurrentRow < len(rs.allData) {
			rs.allData = append(rs.allData[:rs.CurrentRow], rs.allData[rs.CurrentRow+1:]...)
			rs.RecordCount = len(rs.allData)
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
	// Try OLE connection first (for Access databases)
	if conn.oleConnection != nil {
		// Use the stored OLE connection to execute a query and get a recordset
		result, err := oleutil.CallMethod(conn.oleConnection, "Execute", sqlStr)
		if err != nil {
			fmt.Printf("Warning: ADODBRecordset.Open OLE Execute failed: %v\n", err)
			return nil
		}

		if result != nil {
			oleRs := result.ToIDispatch()
			if oleRs != nil {
				// Read all data from the OLE recordset into memory
				rs.State = 1
				rs.allData = make([]map[string]interface{}, 0)

				// Get fields
				fieldsResult, err := oleutil.GetProperty(oleRs, "Fields")
				if err != nil {
					oleRs.Release()
					return nil
				}
				fieldsObj := fieldsResult.ToIDispatch()

				countResult, _ := oleutil.GetProperty(fieldsObj, "Count")
				fieldCount := int(countResult.Val)

				// Get column names
				rs.columns = make([]string, fieldCount)
				for i := 0; i < fieldCount; i++ {
					itemResult, _ := oleutil.GetProperty(fieldsObj, "Item", i)
					field := itemResult.ToIDispatch()
					nameResult, _ := oleutil.GetProperty(field, "Name")
					rs.columns[i] = nameResult.ToString()
					field.Release()
				}

				// Read all rows
				for {
					eofResult, err := oleutil.GetProperty(oleRs, "EOF")
					if err != nil || eofResult.Val != 0 {
						break
					}

					row := make(map[string]interface{})
					for i := 0; i < fieldCount; i++ {
						itemResult, _ := oleutil.GetProperty(fieldsObj, "Item", i)
						field := itemResult.ToIDispatch()
						valueResult, err := oleutil.GetProperty(field, "Value")
						colName := strings.ToLower(rs.columns[i])
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

				fieldsObj.Release()
				oleRs.Release()

				rs.RecordCount = len(rs.allData)

				// Move to first record
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
		}
		return nil
	}

	// Use SQL driver connection
	if conn.db == nil {
		return nil
	}

	rows, err := conn.db.Query(sqlStr)
	if err != nil {
		return nil
	}

	rs.rows = rows
	rs.State = 1
	rs.EOF = false
	rs.BOF = true
	rs.CurrentRow = -1

	// Get column names
	cols, err := rows.Columns()
	if err != nil {
		return nil
	}
	rs.columns = cols

	// Fetch all rows into memory for random access
	rs.allData = make([]map[string]interface{}, 0)
	for rows.Next() {
		// Create slice for values
		values := make([]interface{}, len(cols))
		valuePtrs := make([]interface{}, len(cols))
		for i := range cols {
			valuePtrs[i] = &values[i]
		}

		if err := rows.Scan(valuePtrs...); err != nil {
			continue
		}

		// Convert to map
		row := make(map[string]interface{})
		for i, col := range cols {
			row[strings.ToLower(col)] = values[i]
		}
		rs.allData = append(rs.allData, row)
	}

	rs.RecordCount = len(rs.allData)

	// Move to first record
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

// getRows returns a 2D array containing all records from current position to end
func (rs *ADODBRecordset) getRows(args []interface{}) interface{} {
	if rs.EOF || len(rs.allData) == 0 {
		return [][]interface{}{}
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

	// Sort the data
	sort.Slice(rs.allData, func(i, j int) bool {
		valI := rs.allData[i][rs.SortField]
		valJ := rs.allData[j][rs.SortField]

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

// --- ADODB.Recordset (OLE Wrapper) ---

// ADODBOLERecordset wraps an OLE/COM Recordset object for Access databases
type ADODBOLERecordset struct {
	oleRecordset *ole.IDispatch
	ctx          *ExecutionContext
}

// NewADODBOLERecordset creates a new OLE recordset wrapper
func NewADODBOLERecordset(oleRs *ole.IDispatch, ctx *ExecutionContext) *ADODBOLERecordset {
	return &ADODBOLERecordset{
		oleRecordset: oleRs,
		ctx:          ctx,
	}
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
				return NewADOOLEFields(NewADODBOLEFields(fieldsDisp))
			}else{
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
	}
	return nil
}

func (rs *ADODBOLERecordset) SetProperty(name string, value interface{}) {
	if rs.oleRecordset == nil {
		return
	}

	defer func() {
		if r := recover(); r != nil {
			// Silently recover from OLE errors
		}
	}()

	switch strings.ToLower(name) {
	case "pagesize":
		oleutil.PutProperty(rs.oleRecordset, "PageSize", value)
	case "absolutepage":
		oleutil.PutProperty(rs.oleRecordset, "AbsolutePage", value)
	case "absoluteposition":
		oleutil.PutProperty(rs.oleRecordset, "AbsolutePosition", value)
	}
}

func (rs *ADODBOLERecordset) CallMethod(name string, args ...interface{}) interface{} {
	if rs.oleRecordset == nil {
		return nil
		
	}

	defer func() {
		if r := recover(); r != nil {
		}
	}()

	method := strings.ToLower(name)

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
				// Get the field by name/index
				fieldResult, err := oleutil.GetProperty(fieldsDisp, "Item", args[0])
				if err != nil {
					fmt.Printf("Error> ADODB.Recordset CallMethod Fields Item error: %s\n", err.Error())
					return nil
				}
				fieldDisp := fieldResult.ToIDispatch()
				if fieldDisp != nil {
					// Get the field value
					valueResult, err := oleutil.GetProperty(fieldDisp, "Value")
					if err != nil {
						fmt.Printf("Error> ADODB.Recordset CallMethod Fields Item Value error: %s\n", err.Error())
						return nil
					}
					return valueResult.Value()
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
			oleutil.CallMethod(rs.oleRecordset, "Open", args...)
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
		oleutil.CallMethod(rs.oleRecordset, "AddNew")
		return nil
	case "update":
		if len(args) > 0 {
			oleutil.CallMethod(rs.oleRecordset, "Update", args...)
		} else {
			oleutil.CallMethod(rs.oleRecordset, "Update")
		}
		return nil
	case "delete":
		oleutil.CallMethod(rs.oleRecordset, "Delete")
		return nil
	case "cancelupdate":
		oleutil.CallMethod(rs.oleRecordset, "CancelUpdate")
		return nil
	}

	return nil
}

// --- ADODB.Fields (OLE Wrapper) ---

// ADODBOLEFields wraps an OLE/COM Fields collection
type ADODBOLEFields struct {
	oleFields *ole.IDispatch
}

// NewADODBOLEFields creates a new OLE fields wrapper
func NewADODBOLEFields(oleFields *ole.IDispatch) *ADODBOLEFields {
	return &ADODBOLEFields{
		oleFields: oleFields,
	}
}

func (f *ADODBOLEFields) GetProperty(name string) interface{} {
	if f.oleFields == nil {
		return nil
	}

	defer func() {
		if r := recover(); r != nil {
			// Silently recover from OLE errors
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
			valueResult, err := oleutil.GetProperty(fieldDisp, "Value")
			if err != nil {
				return nil
			}
			return valueResult.Value()
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
				// Get field value
				valueResult, err := oleutil.GetProperty(fieldDisp, "Value")
				if err != nil {
					return nil
				}
				return valueResult.Value()
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
