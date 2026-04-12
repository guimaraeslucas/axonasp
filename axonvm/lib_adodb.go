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
	"database/sql"
	"fmt"
	"math"
	"net/url"
	"os"
	"path/filepath"
	"regexp"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
	"unicode"
	"unicode/utf16"

	"g3pix.com.br/axonasp/vbscript"
	_ "github.com/denisenkom/go-mssqldb"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	_ "github.com/go-sql-driver/mysql"
	_ "github.com/lib/pq"
	_ "github.com/sijms/go-ora/v2"
	_ "modernc.org/sqlite"
)

// ADODB Constants
const (
	adStateClosed = 0
	adStateOpen   = 1

	adModeRead          = 1
	adModeWrite         = 2
	adModeReadWrite     = 3
	adModeShareDenyRead = 4
	adModeShareDenyNone = 16

	adOpenForwardOnly = 0
	adOpenKeyset      = 1
	adOpenDynamic     = 2
	adOpenStatic      = 3

	adLockReadOnly        = 1
	adLockPessimistic     = 2
	adLockOptimistic      = 3
	adLockBatchOptimistic = 4

	adUseServer = 2
	adUseClient = 3

	adSchemaProcedures  = 16
	adSchemaViews       = 23
	adSchemaForeignKeys = 27
	adSchemaColumns     = 4
	adSchemaIndexes     = 12

	adEditNone       = 0
	adEditInProgress = 1
	adEditAdd        = 2

	adRecOK = 0
)

var (
	adodbPlatformArchitectureOnce         sync.Once
	adodbPlatformArchitectureCached       = "auto"
	adodbPlatformArchitectureTestOverride string
)

// adodbConnection stores one ADODB.Connection runtime instance.
type adodbConnection struct {
	connectionString  string
	provider          string
	version           string
	commandTimeout    int
	connectionTimeout int
	cursorLocation    int
	defaultDatabase   string
	isolationLevel    int
	state             int
	mode              int
	errors            []adodbError
	db                *sql.DB
	dbDriver          string
	tx                *sql.Tx
	oleConnection     *ole.IDispatch
	lastHealthCheck   time.Time
}

// adodbRecordset stores one ADODB.Recordset runtime instance.
type adodbRecordset struct {
	eof                 bool
	bof                 bool
	state               int
	currentRow          int
	bookmark            int
	recordCount         int
	pageSize            int
	absolutePage        int
	cacheSize           int
	filter              string
	sort                string
	dataMember          string
	editMode            int
	index               string
	marshalOptions      int
	maxRecords          int
	status              int
	activeCommand       int64
	activeConnection    int64 // objID of adodbConnection
	cursorType          int
	lockType            int
	cursorLocation      int
	source              string
	columns             []string
	columnIndexByLower  map[string]int
	columnTypes         []string
	columnTypeByName    map[string]int
	columnSizeByName    map[string]int
	columnAttrByName    map[string]int
	columnScaleByName   map[string]int
	fieldChunkOffset    map[string]int
	pendingUpdateFields map[string]struct{}
	data                []map[string]Value
	sqlRows             *sql.Rows
	oleRecordset        *ole.IDispatch
}

// adodbCommand stores one ADODB.Command runtime instance.
type adodbCommand struct {
	activeConnection int64 // objID of adodbConnection
	commandText      string
	commandType      int
	commandTimeout   int
	prepared         bool
	parameters       []*adodbParameter
}

// adodbParameter stores one ADODB.Parameter runtime instance.
type adodbParameter struct {
	name      string
	typ       int
	direction int
	size      int
	value     Value
}

// adodbError stores one ADODB.Error runtime instance.
type adodbError struct {
	number      int
	description string
	source      string
	sqlState    string
}

// newADODBConnection allocates one ADODB.Connection native object.
func (vm *VM) newADODBConnection() Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.adodbConnectionItems[objID] = &adodbConnection{
		version:           "6.1",
		commandTimeout:    30,
		connectionTimeout: 15,
		cursorLocation:    adUseServer,
		state:             adStateClosed,
		mode:              adModeReadWrite,
		errors:            make([]adodbError, 0, 2),
	}
	return Value{Type: VTNativeObject, Num: objID}
}

// newADODBOLEConnection allocates one ADODBOLE.Connection native object (Access alias).
func (vm *VM) newADODBOLEConnection() Value {
	// For AxonASP, we use the same structure, but the user requested aliases.
	// Internal logic handles the connection string.
	return vm.newADODBConnection()
}

// newADODBRecordset allocates one ADODB.Recordset native object.
func (vm *VM) newADODBRecordset() Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.adodbRecordsetItems[objID] = &adodbRecordset{
		state:               adStateClosed,
		eof:                 true,
		bof:                 true,
		currentRow:          -1,
		bookmark:            0,
		pageSize:            10,
		absolutePage:        1,
		cacheSize:           1,
		editMode:            adEditNone,
		status:              adRecOK,
		columnIndexByLower:  make(map[string]int),
		columnTypeByName:    make(map[string]int),
		columnSizeByName:    make(map[string]int),
		columnAttrByName:    make(map[string]int),
		columnScaleByName:   make(map[string]int),
		fieldChunkOffset:    make(map[string]int),
		pendingUpdateFields: make(map[string]struct{}),
	}
	return Value{Type: VTNativeObject, Num: objID}
}

// newADODBCommand allocates one ADODB.Command native object.
func (vm *VM) newADODBCommand() Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.adodbCommandItems[objID] = &adodbCommand{
		commandType:    1,
		commandTimeout: 30,
		parameters:     make([]*adodbParameter, 0),
	}
	return Value{Type: VTNativeObject, Num: objID}
}

// dispatchADODBMethod routes ADODB method calls.
func (vm *VM) dispatchADODBMethod(objID int64, member string, args []Value) (Value, bool) {
	if conn, exists := vm.adodbConnectionItems[objID]; exists {
		return vm.dispatchADODBConnectionMethod(conn, member, args), true
	}
	if rs, exists := vm.adodbRecordsetItems[objID]; exists {
		return vm.dispatchADODBRecordsetMethod(rs, member, args), true
	}
	if cmd, exists := vm.adodbCommandItems[objID]; exists {
		return vm.dispatchADODBCommandMethod(cmd, member, args), true
	}
	return Value{Type: VTEmpty}, false
}

// dispatchADODBPropertyGet resolves ADODB property reads.
func (vm *VM) dispatchADODBPropertyGet(objID int64, member string) (Value, bool) {
	if conn, exists := vm.adodbConnectionItems[objID]; exists {
		return vm.dispatchADODBConnectionPropertyGet(conn, member), true
	}
	if rs, exists := vm.adodbRecordsetItems[objID]; exists {
		return vm.dispatchADODBRecordsetPropertyGet(rs, member), true
	}
	if cmd, exists := vm.adodbCommandItems[objID]; exists {
		return vm.dispatchADODBCommandPropertyGet(cmd, member), true
	}
	return Value{Type: VTEmpty}, false
}

// dispatchADODBPropertySet handles ADODB writable properties.
func (vm *VM) dispatchADODBPropertySet(objID int64, member string, val Value) bool {
	if conn, exists := vm.adodbConnectionItems[objID]; exists {
		return vm.dispatchADODBConnectionPropertySet(conn, member, val)
	}
	if rs, exists := vm.adodbRecordsetItems[objID]; exists {
		return vm.dispatchADODBRecordsetPropertySet(rs, member, val)
	}
	if cmd, exists := vm.adodbCommandItems[objID]; exists {
		return vm.dispatchADODBCommandPropertySet(cmd, member, val)
	}
	return false
}

// --- Connection Implementation ---

func (vm *VM) dispatchADODBConnectionMethod(conn *adodbConnection, member string, args []Value) Value {
	switch {
	case strings.EqualFold(member, "Open"):
		if len(args) > 0 {
			conn.connectionString = args[0].String()
		}
		vm.adodbConnectionOpen(conn)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Close"):
		vm.adodbConnectionClose(conn)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Execute"):
		return vm.adodbConnectionExecute(conn, args)
	case strings.EqualFold(member, "BeginTrans"):
		return vm.adodbConnectionBeginTrans(conn)
	case strings.EqualFold(member, "CommitTrans"):
		vm.adodbConnectionCommitTrans(conn)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "RollbackTrans"):
		vm.adodbConnectionRollbackTrans(conn)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "OpenSchema"):
		return vm.adodbConnectionOpenSchema(conn, args)
	case strings.EqualFold(member, "Cancel"):
		return Value{Type: VTEmpty}
	}
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchADODBConnectionPropertyGet(conn *adodbConnection, member string) Value {
	switch {
	case strings.EqualFold(member, "ConnectionString"):
		return NewString(conn.connectionString)
	case strings.EqualFold(member, "State"):
		return NewInteger(int64(conn.state))
	case strings.EqualFold(member, "Mode"):
		return NewInteger(int64(conn.mode))
	case strings.EqualFold(member, "Provider"):
		return NewString(conn.provider)
	case strings.EqualFold(member, "Version"):
		return NewString(conn.version)
	case strings.EqualFold(member, "CommandTimeout"):
		return NewInteger(int64(conn.commandTimeout))
	case strings.EqualFold(member, "ConnectionTimeout"):
		return NewInteger(int64(conn.connectionTimeout))
	case strings.EqualFold(member, "CursorLocation"):
		return NewInteger(int64(conn.cursorLocation))
	case strings.EqualFold(member, "DefaultDatabase"):
		return NewString(conn.defaultDatabase)
	case strings.EqualFold(member, "IsolationLevel"):
		return NewInteger(int64(conn.isolationLevel))
	case strings.EqualFold(member, "Errors"):
		// Return a simulated Errors collection
		return vm.newADODBErrorsCollection(conn)
	}
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchADODBConnectionPropertySet(conn *adodbConnection, member string, val Value) bool {
	switch {
	case strings.EqualFold(member, "ConnectionString"):
		conn.connectionString = val.String()
		return true
	case strings.EqualFold(member, "Mode"):
		conn.mode = vm.asInt(val)
		return true
	case strings.EqualFold(member, "Provider"):
		conn.provider = strings.TrimSpace(val.String())
		return true
	case strings.EqualFold(member, "CommandTimeout"):
		conn.commandTimeout = vm.asInt(val)
		return true
	case strings.EqualFold(member, "ConnectionTimeout"):
		conn.connectionTimeout = vm.asInt(val)
		return true
	case strings.EqualFold(member, "CursorLocation"):
		conn.cursorLocation = vm.asInt(val)
		return true
	case strings.EqualFold(member, "DefaultDatabase"):
		conn.defaultDatabase = strings.TrimSpace(val.String())
		return true
	case strings.EqualFold(member, "IsolationLevel"):
		conn.isolationLevel = vm.asInt(val)
		return true
	}
	return false
}

func (vm *VM) adodbConnectionOpen(conn *adodbConnection) {
	if conn.state == adStateOpen {
		vm.adodbConnectionClearErrors(conn)
		return
	}
	vm.adodbConnectionClearErrors(conn)

	connStr := strings.TrimSpace(conn.connectionString)
	if connStr == "" {
		vm.adodbConnectionRaiseProviderError(conn, "ADODB.Connection.Open", "connection string is empty", "")
		return
	}
	if runtime.GOOS == "windows" {
		connStr = adodbApplyPlatformAccessProvider(connStr)
		conn.connectionString = connStr
	}

	vm.adodbApplyConnectionMetadata(conn, connStr)

	// Windows OLE/Access check
	connStrLower := strings.ToLower(connStr)
	if (strings.Contains(connStrLower, "microsoft.jet.oledb") || strings.Contains(connStrLower, "microsoft.ace.oledb")) && runtime.GOOS == "windows" {
		vm.adodbOpenAccessDatabase(conn, connStr)
		return
	}

	driver, dsn := vm.adodbParseConnectionString(connStr)
	if driver == "" {
		vm.adodbConnectionRaiseProviderError(conn, "ADODB.Connection", "unsupported connection string", "")
		return
	}

	db, err := sql.Open(driver, dsn)
	if err != nil {
		vm.adodbConnectionRaiseProviderError(conn, "ADODB.Connection", err.Error(), "")
		return
	}

	if err := db.Ping(); err != nil {
		db.Close()
		vm.adodbConnectionRaiseProviderError(conn, "ADODB.Connection", "ping failed: "+err.Error(), "")
		return
	}

	// SQLite optimization
	if driver == "sqlite" {
		// Keep one physical connection for ADODB-like behavior and to avoid
		// cross-connection visibility issues (notably with :memory: databases).
		db.SetMaxOpenConns(1)
		db.SetMaxIdleConns(1)
		_, _ = db.Exec("PRAGMA journal_mode = WAL")
		_, _ = db.Exec("PRAGMA busy_timeout = 5000")
	}

	conn.db = db
	conn.dbDriver = driver
	conn.state = adStateOpen
}

// adodbConnectionOpenSchema returns a lightweight schema recordset for supported providers.
func (vm *VM) adodbConnectionOpenSchema(conn *adodbConnection, args []Value) Value {
	if conn == nil {
		return Value{Type: VTEmpty}
	}
	if conn.state != adStateOpen {
		vm.adodbConnectionOpen(conn)
	}
	schemaID := adSchemaTables
	if len(args) >= 1 {
		schemaID = vm.asInt(args[0])
	}
	restrictions := vm.adodbSchemaRestrictions(args)
	rsVal := vm.newADODBRecordset()
	rs := vm.adodbRecordsetItems[rsVal.Num]
	if rs == nil {
		return Value{Type: VTEmpty}
	}
	rs.columns = vm.adodbSchemaColumnNames(schemaID)
	vm.adodbRecordsetRebuildColumnIndex(rs)
	rs.source = "__schema__"
	rs.activeConnection = 0
	rs.state = adStateOpen
	rs.data = make([]map[string]Value, 0, 8)

	if conn.db != nil {
		if rows := vm.adodbBuildSchemaRows(conn, schemaID, restrictions); len(rows) > 0 {
			rs.data = append(rs.data, rows...)
		}
	}

	if conn.oleConnection != nil && len(rs.data) == 0 {
		res, err := vm.adodbConnectionOpenSchemaOLE(conn, schemaID, restrictions)
		if err == nil && res != nil {
			disp := res.ToIDispatch()
			if disp != nil {
				disp.AddRef()
				rs.oleRecordset = disp
				vm.adodbPopulateRecordsetFromOLE(rs)
			}
			res.Clear()
		}
	}

	vm.adodbFinalizeDisconnectedRecordset(rs)
	return rsVal
}

// adodbConnectionOpenSchemaOLE forwards OpenSchema calls to OLE providers and preserves restrictions when present.
func (vm *VM) adodbConnectionOpenSchemaOLE(conn *adodbConnection, schemaID int, restrictions []string) (*ole.VARIANT, error) {
	if conn == nil || conn.oleConnection == nil {
		return nil, nil
	}
	if !adodbHasNonEmptyRestriction(restrictions) {
		return oleutil.CallMethod(conn.oleConnection, "OpenSchema", int32(schemaID))
	}
	restrictionArgs := make([]string, len(restrictions))
	copy(restrictionArgs, restrictions)
	return oleutil.CallMethod(conn.oleConnection, "OpenSchema", int32(schemaID), restrictionArgs)
}

func (vm *VM) adodbConnectionClose(conn *adodbConnection) {
	if conn.tx != nil {
		_ = conn.tx.Rollback()
		conn.tx = nil
	}
	if conn.db != nil {
		_ = conn.db.Close()
		conn.db = nil
	}
	if conn.oleConnection != nil {
		res, _ := oleutil.CallMethod(conn.oleConnection, "Close")
		if res != nil {
			res.Clear()
		}
		conn.oleConnection.Release()
		conn.oleConnection = nil
	}
	conn.state = adStateClosed
}

// CleanupRequestResources deterministically releases native ADODB resources owned by one VM.
// This prevents leaked OLE connections and locked OS threads when ASP code forgets explicit Close.
func (vm *VM) CleanupRequestResources() {
	if vm == nil {
		return
	}

	// Close recordsets first so any provider cursors are released before connection shutdown.
	for _, rs := range vm.adodbRecordsetItems {
		if rs != nil {
			vm.adodbRecordsetClose(rs)
		}
	}

	for _, conn := range vm.adodbConnectionItems {
		if conn != nil {
			vm.adodbConnectionClose(conn)
		}
	}

	// Clear dynamic native-object maps to release references promptly.
	vm.adodbRecordsetItems = make(map[int64]*adodbRecordset)
	vm.adodbConnectionItems = make(map[int64]*adodbConnection)
	vm.adodbCommandItems = make(map[int64]*adodbCommand)
	vm.adodbParameterItems = make(map[int64]*adodbParameter)
	vm.adodbErrorsCollectionItems = make(map[int64]*adodbConnection)
	vm.adodbErrorItems = make(map[int64]*adodbError)
	vm.adodbFieldsCollectionItems = make(map[int64]*adodbRecordset)
	vm.adodbParametersCollectionItems = make(map[int64]*adodbCommand)
	vm.adodbFieldItems = make(map[int64]*adodbFieldProxy)
	// Clear the generic native-object proxy map to prevent cross-request leaks.
	vm.nativeObjectProxies = make(map[int64]nativeObjectProxy)
	vm.nextDynamicNativeID = 20000
	vm.releaseCOMRequestThread()
}

func (vm *VM) ensureCOMRequestThread() error {
	if vm == nil || vm.comThreadLocked {
		return nil
	}
	runtime.LockOSThread()
	initialized, err := comInitialize()
	if err != nil {
		runtime.UnlockOSThread()
		return err
	}
	vm.comThreadLocked = true
	vm.comInitialized = initialized
	return nil
}

func (vm *VM) releaseCOMRequestThread() {
	if vm == nil {
		return
	}
	if vm.comInitialized {
		ole.CoUninitialize()
		vm.comInitialized = false
	}
	if vm.comThreadLocked {
		runtime.UnlockOSThread()
		vm.comThreadLocked = false
	}
}

func (vm *VM) adodbConnectionExecute(conn *adodbConnection, args []Value) Value {
	vm.adodbConnectionClearErrors(conn)
	if conn.state != adStateOpen {
		vm.adodbConnectionOpen(conn)
		if conn.state != adStateOpen {
			return Value{Type: VTEmpty}
		}
	}

	if len(args) == 0 {
		return Value{Type: VTEmpty}
	}

	sqlText := args[0].String()
	isQuery := vm.adodbIsQuery(sqlText)
	execArgs := vm.adodbExecuteArgs(args)

	if conn.db != nil {
		if isQuery {
			if len(execArgs) == 0 {
				rsVal := vm.newADODBRecordset()
				rs := vm.adodbRecordsetItems[rsVal.Num]
				vm.adodbRecordsetOpen(rs, sqlText, conn, args)
				return rsVal
			}

			// Parameterized query path for ADODB.Connection.Execute sql, params.
			var (
				rows *sql.Rows
				err  error
			)
			if conn.tx != nil {
				rows, err = conn.tx.Query(sqlText, execArgs...)
			} else {
				rows, err = conn.db.Query(sqlText, execArgs...)
			}
			if err != nil {
				vm.adodbConnectionRaiseProviderError(conn, "ADODB.Connection.Execute", err.Error(), "")
				return Value{Type: VTEmpty}
			}
			rsVal := vm.newADODBRecordset()
			rs := vm.adodbRecordsetItems[rsVal.Num]
			rs.sqlRows = rows
			vm.adodbRecordsetLoadCurrentSQLResultSet(rs)
			return rsVal
		}

		var (
			res sql.Result
			err error
		)
		if conn.tx != nil {
			res, err = conn.tx.Exec(sqlText, execArgs...)
		} else {
			res, err = conn.db.Exec(sqlText, execArgs...)
		}
		if err != nil {
			vm.adodbConnectionRaiseProviderError(conn, "ADODB.Connection.Execute", err.Error(), "")
			return Value{Type: VTEmpty}
		}
		affected, _ := res.RowsAffected()
		return NewInteger(affected)
	}

	if conn.oleConnection != nil {
		// OLE Execute
		res, err := oleutil.CallMethod(conn.oleConnection, "Execute", sqlText)
		if err != nil {
			vm.adodbConnectionRaiseProviderError(conn, "ADODB.Connection.Execute", "OLE: "+err.Error(), "")
			return Value{Type: VTEmpty}
		}
		if res != nil {
			defer res.Clear()
			if isQuery {
				disp := res.ToIDispatch()
				if disp != nil {
					disp.AddRef()
					rsVal := vm.newADODBRecordset()
					rs := vm.adodbRecordsetItems[rsVal.Num]
					rs.oleRecordset = disp
					rs.state = adStateOpen
					vm.adodbPopulateRecordsetFromOLE(rs)
					return rsVal
				}
			}
		}
	}

	return Value{Type: VTEmpty}
}

func (vm *VM) adodbConnectionBeginTrans(conn *adodbConnection) Value {
	vm.adodbConnectionClearErrors(conn)
	if conn.db == nil {
		return NewInteger(0)
	}
	tx, err := conn.db.Begin()
	if err != nil {
		vm.adodbConnectionRaiseProviderError(conn, "ADODB.Connection.BeginTrans", err.Error(), "")
		return NewInteger(0)
	}
	conn.tx = tx
	return NewInteger(1)
}

func (vm *VM) adodbConnectionCommitTrans(conn *adodbConnection) {
	vm.adodbConnectionClearErrors(conn)
	if conn.tx == nil {
		return
	}
	err := conn.tx.Commit()
	conn.tx = nil
	if err != nil {
		vm.adodbConnectionRaiseProviderError(conn, "ADODB.Connection.CommitTrans", err.Error(), "")
	}
}

func (vm *VM) adodbConnectionRollbackTrans(conn *adodbConnection) {
	if conn.tx == nil {
		return
	}
	_ = conn.tx.Rollback()
	conn.tx = nil
}

// --- Recordset Implementation ---

func (vm *VM) dispatchADODBRecordsetMethod(rs *adodbRecordset, member string, args []Value) Value {
	switch {
	case strings.EqualFold(member, "Fields.Append"):
		if len(args) > 0 {
			name := args[0].String()
			fieldType := 0
			definedSize := 0
			attributes := 0
			if len(args) > 1 && args[1].Type != VTEmpty {
				fieldType = vm.asInt(args[1])
			}
			if len(args) > 2 && args[2].Type != VTEmpty {
				definedSize = vm.asInt(args[2])
			}
			if len(args) > 3 && args[3].Type != VTEmpty {
				attributes = vm.asInt(args[3])
			}
			vm.adodbRecordsetAppendField(rs, name, fieldType, definedSize, attributes)
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Fields.Item"):
		if len(args) > 0 {
			return vm.newADODBFieldProxy(rs, args[0].String())
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Open"):
		vm.adodbRecordsetOpen(rs, "", nil, args)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Close"):
		vm.adodbRecordsetClose(rs)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "MoveNext"):
		vm.adodbRecordsetMoveNext(rs)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "MovePrevious"):
		vm.adodbRecordsetMovePrevious(rs)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "MoveFirst"):
		vm.adodbRecordsetMoveFirst(rs)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "MoveLast"):
		vm.adodbRecordsetMoveLast(rs)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Move"):
		vm.adodbRecordsetMove(rs, args)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Cancel"):
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "CancelBatch"):
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "CancelUpdate"):
		rs.editMode = adEditNone
		vm.adodbRecordsetClearPendingUpdateFields(rs)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Find"):
		vm.adodbRecordsetFind(rs, args)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "NextRecordset"):
		return vm.adodbRecordsetNextRecordset(rs)
	case strings.EqualFold(member, "Clone"):
		return vm.adodbRecordsetClone(rs)
	case strings.EqualFold(member, "CompareBookmarks"):
		return vm.adodbRecordsetCompareBookmarks(args)
	case strings.EqualFold(member, "AddNew"):
		vm.adodbRecordsetAddNew(rs, args)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Update"):
		vm.adodbRecordsetUpdate(rs, args)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Delete"):
		vm.adodbRecordsetDelete(rs)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "GetRows"):
		return vm.adodbRecordsetGetRows(rs, args)
	case strings.EqualFold(member, "GetString"):
		return vm.adodbRecordsetGetString(rs, args)
	case strings.EqualFold(member, "Requery"):
		vm.adodbRecordsetRequery(rs)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Resync"):
		vm.adodbRecordsetResync(rs)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Save"):
		vm.adodbRecordsetSave(rs, args)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Seek"):
		vm.adodbRecordsetSeek(rs, args)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Supports"):
		return NewBool(vm.adodbRecordsetSupports(rs, args))
	case strings.EqualFold(member, "UpdateBatch"):
		vm.adodbRecordsetUpdateBatch(rs)
		return Value{Type: VTEmpty}
	case member == "":
		if len(args) > 0 {
			if vm.adodbRecordsetIsLiveOLE(rs) {
				key := vm.adodbRecordsetResolveColumnKey(rs, args[0])
				if idx, exists := rs.columnIndexByLower[key]; exists && idx >= 0 && idx < len(rs.columns) {
					key = rs.columns[idx]
				}
				if key != "" {
					if len(args) > 1 {
						if vm.adodbOLESetFieldValue(rs, key, args[len(args)-1]) {
							if rs.editMode != adEditAdd {
								rs.editMode = adEditInProgress
							}
							vm.adodbRecordsetMarkPendingUpdateField(rs, key)
						}
						return Value{Type: VTEmpty}
					}
					if v, ok := vm.adodbOLEGetFieldValue(rs, key); ok {
						return v
					}
				}
			}
			if rs.state == adStateOpen && rs.currentRow >= 0 && rs.currentRow < len(rs.data) {
				row := rs.data[rs.currentRow]
				if row != nil {
					key := vm.adodbRecordsetResolveColumnKey(rs, args[0])
					if key != "" {
						if len(args) > 1 {
							row[key] = args[len(args)-1]
							if rs.editMode != adEditAdd {
								rs.editMode = adEditInProgress
							}
							vm.adodbRecordsetMarkPendingUpdateField(rs, key)
							return Value{Type: VTEmpty}
						}
						if val, ok := row[key]; ok {
							return val
						}
					}
				}
			}
		}
	case strings.EqualFold(member, "Fields"):
		if len(args) == 0 {
			return vm.newADODBFieldsCollection(rs)
		}
		if args[0].Type == VTInteger || args[0].Type == VTDouble {
			idx := vm.asInt(args[0])
			if idx >= 0 && idx < len(rs.columns) {
				return vm.newADODBFieldProxy(rs, rs.columns[idx])
			}
			return Value{Type: VTEmpty}
		}
		name := strings.TrimSpace(args[0].String())
		if name == "" {
			return Value{Type: VTEmpty}
		}
		return vm.newADODBFieldProxy(rs, name)
	}
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchADODBRecordsetPropertyGet(rs *adodbRecordset, member string) Value {
	switch {
	case strings.EqualFold(member, "EOF"):
		if vm.adodbRecordsetIsLiveOLE(rs) {
			vm.adodbOLERefreshCursorFlags(rs)
		}
		return NewBool(rs.eof)
	case strings.EqualFold(member, "BOF"):
		if vm.adodbRecordsetIsLiveOLE(rs) {
			vm.adodbOLERefreshCursorFlags(rs)
		}
		return NewBool(rs.bof)
	case strings.EqualFold(member, "State"):
		return NewInteger(int64(rs.state))
	case strings.EqualFold(member, "RecordCount"):
		if vm.adodbRecordsetIsLiveOLE(rs) {
			countRes, _ := oleutil.GetProperty(rs.oleRecordset, "RecordCount")
			if countRes != nil {
				rs.recordCount = vm.asInt(vm.adodbValueToVMValue(countRes.Value()))
				countRes.Clear()
			}
		}
		return NewInteger(int64(rs.recordCount))
	case strings.EqualFold(member, "Fields"):
		return vm.newADODBFieldsCollection(rs)
	case strings.EqualFold(member, "AbsolutePage"):
		return NewInteger(int64(rs.absolutePage))
	case strings.EqualFold(member, "PageCount"):
		if rs.pageSize <= 0 {
			return NewInteger(0)
		}
		count := (rs.recordCount + rs.pageSize - 1) / rs.pageSize
		return NewInteger(int64(count))
	case strings.EqualFold(member, "PageSize"):
		return NewInteger(int64(rs.pageSize))
	case strings.EqualFold(member, "AbsolutePosition"):
		if vm.adodbRecordsetIsLiveOLE(rs) {
			posRes, _ := oleutil.GetProperty(rs.oleRecordset, "AbsolutePosition")
			if posRes != nil {
				pos := vm.asInt(vm.adodbValueToVMValue(posRes.Value()))
				if pos > 0 {
					rs.currentRow = pos - 1
					rs.bookmark = pos
				}
				posRes.Clear()
			}
		}
		if rs.currentRow < 0 {
			return NewInteger(0)
		}
		return NewInteger(int64(rs.currentRow + 1))
	case strings.EqualFold(member, "Bookmark"):
		if rs.currentRow < 0 {
			return NewInteger(0)
		}
		return NewInteger(int64(rs.currentRow + 1))
	case strings.EqualFold(member, "ActiveCommand"):
		if rs.activeCommand == 0 {
			return Value{Type: VTEmpty}
		}
		return Value{Type: VTNativeObject, Num: rs.activeCommand}
	case strings.EqualFold(member, "CacheSize"):
		return NewInteger(int64(rs.cacheSize))
	case strings.EqualFold(member, "DataMember"):
		return NewString(rs.dataMember)
	case strings.EqualFold(member, "EditMode"):
		return NewInteger(int64(rs.editMode))
	case strings.EqualFold(member, "Filter"):
		return NewString(rs.filter)
	case strings.EqualFold(member, "Index"):
		return NewString(rs.index)
	case strings.EqualFold(member, "MarshalOptions"):
		return NewInteger(int64(rs.marshalOptions))
	case strings.EqualFold(member, "MaxRecords"):
		return NewInteger(int64(rs.maxRecords))
	case strings.EqualFold(member, "Sort"):
		return NewString(rs.sort)
	case strings.EqualFold(member, "Source"):
		return NewString(rs.source)
	case strings.EqualFold(member, "Status"):
		return NewInteger(int64(rs.status))
	default:
		// Direct field access: rs("FieldName")
		if vm.adodbRecordsetIsLiveOLE(rs) {
			if val, ok := vm.adodbOLEGetFieldValue(rs, member); ok {
				return val
			}
		}
		if rs.state == adStateOpen && rs.currentRow >= 0 && rs.currentRow < len(rs.data) {
			row := rs.data[rs.currentRow]
			key := vm.adodbRecordsetResolveColumnKey(rs, NewString(member))
			if val, ok := row[key]; ok {
				return val
			}
		}
	}
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchADODBRecordsetPropertySet(rs *adodbRecordset, member string, val Value) bool {
	switch {
	case strings.EqualFold(member, "ActiveConnection"):
		if val.Type == VTNativeObject {
			rs.activeConnection = val.Num
		}
		return true
	case strings.EqualFold(member, "ActiveCommand"):
		if val.Type == VTNativeObject {
			rs.activeCommand = val.Num
			if cmd, ok := vm.adodbCommandItems[val.Num]; ok {
				rs.source = cmd.commandText
				if cmd.activeConnection != 0 {
					rs.activeConnection = cmd.activeConnection
				}
			}
		}
		return true
	case strings.EqualFold(member, "CursorType"):
		rs.cursorType = vm.asInt(val)
		return true
	case strings.EqualFold(member, "LockType"):
		rs.lockType = vm.asInt(val)
		return true
	case strings.EqualFold(member, "CursorLocation"):
		rs.cursorLocation = vm.asInt(val)
		return true
	case strings.EqualFold(member, "PageSize"):
		rs.pageSize = vm.asInt(val)
		if vm.adodbRecordsetIsLiveOLE(rs) {
			_, _ = oleutil.PutProperty(rs.oleRecordset, "PageSize", int32(rs.pageSize))
		}
		return true
	case strings.EqualFold(member, "AbsolutePage"):
		rs.absolutePage = vm.asInt(val)
		return true
	case strings.EqualFold(member, "AbsolutePosition"):
		pos := vm.asInt(val)
		if vm.adodbRecordsetIsLiveOLE(rs) {
			_, _ = oleutil.PutProperty(rs.oleRecordset, "AbsolutePosition", int32(pos))
			if pos > 0 {
				rs.currentRow = pos - 1
				rs.bookmark = pos
			}
			vm.adodbOLERefreshCursorFlags(rs)
			return true
		}
		if pos <= 0 {
			rs.currentRow = -1
			rs.bof = true
			return true
		}
		idx := pos - 1
		if idx >= rs.recordCount {
			rs.currentRow = rs.recordCount
			rs.eof = true
			return true
		}
		rs.currentRow = idx
		rs.bof = false
		rs.eof = false
		rs.bookmark = pos
		return true
	case strings.EqualFold(member, "Bookmark"):
		pos := vm.asInt(val)
		if vm.adodbRecordsetIsLiveOLE(rs) {
			_, _ = oleutil.PutProperty(rs.oleRecordset, "Bookmark", int32(pos))
			if pos > 0 {
				rs.currentRow = pos - 1
				rs.bookmark = pos
			}
			vm.adodbOLERefreshCursorFlags(rs)
			return true
		}
		if pos > 0 && pos <= rs.recordCount {
			rs.currentRow = pos - 1
			rs.bookmark = pos
			rs.eof = false
			rs.bof = false
		}
		return true
	case strings.EqualFold(member, "CacheSize"):
		size := vm.asInt(val)
		if size <= 0 {
			size = 1
		}
		rs.cacheSize = size
		return true
	case strings.EqualFold(member, "Filter"):
		rs.filter = strings.TrimSpace(vm.valueToString(val))
		vm.adodbRecordsetApplyFilterSort(rs)
		return true
	case strings.EqualFold(member, "DataMember"):
		rs.dataMember = strings.TrimSpace(vm.valueToString(val))
		return true
	case strings.EqualFold(member, "Index"):
		rs.index = strings.TrimSpace(vm.valueToString(val))
		return true
	case strings.EqualFold(member, "MarshalOptions"):
		rs.marshalOptions = vm.asInt(val)
		return true
	case strings.EqualFold(member, "MaxRecords"):
		rs.maxRecords = vm.asInt(val)
		return true
	case strings.EqualFold(member, "Sort"):
		rs.sort = strings.TrimSpace(vm.valueToString(val))
		vm.adodbRecordsetApplyFilterSort(rs)
		return true
	case strings.EqualFold(member, "Source"):
		if val.Type == VTNativeObject {
			rs.activeCommand = val.Num
			if cmd, ok := vm.adodbCommandItems[val.Num]; ok {
				rs.source = cmd.commandText
				if cmd.activeConnection != 0 {
					rs.activeConnection = cmd.activeConnection
				}
			}
		} else {
			rs.source = strings.TrimSpace(vm.valueToString(val))
		}
		return true
	case strings.EqualFold(member, "Status"):
		rs.status = vm.asInt(val)
		return true
	}
	return false
}

func (vm *VM) adodbRecordsetOpen(rs *adodbRecordset, sqlText string, conn *adodbConnection, args []Value) {
	if conn != nil {
		vm.adodbConnectionClearErrors(conn)
	}
	if rs.state == adStateOpen {
		vm.adodbRecordsetClose(rs)
	}

	if sqlText == "" && len(args) > 0 {
		sqlText = args[0].String()
	}
	if conn == nil && len(args) > 1 {
		if args[1].Type == VTNativeObject {
			conn = vm.adodbConnectionItems[args[1].Num]
		}
	}
	if conn == nil && rs.activeConnection != 0 {
		conn = vm.adodbConnectionItems[rs.activeConnection]
	}
	if conn != nil {
		rs.activeConnection = 0
		for objID, current := range vm.adodbConnectionItems {
			if current == conn {
				rs.activeConnection = objID
				break
			}
		}
	}
	if strings.TrimSpace(sqlText) != "" {
		rs.source = sqlText
	}

	if conn == nil && strings.TrimSpace(sqlText) == "" && len(rs.columns) > 0 {
		if rs.data == nil {
			rs.data = make([]map[string]Value, 0)
		}
		rs.recordCount = len(rs.data)
		rs.state = adStateOpen
		if rs.recordCount > 0 {
			rs.currentRow = 0
			rs.eof = false
			rs.bof = false
		} else {
			rs.currentRow = -1
			rs.eof = true
			rs.bof = true
		}
		return
	}

	if conn == nil {
		vm.raise(vbscript.InternalError,
			"ADODB.Recordset.Open: no active connection")
		return
	}

	if conn.state != adStateOpen {
		vm.adodbConnectionOpen(conn)
	}
	sqlText = vm.adodbNormalizeRecordsetSource(sqlText, conn)

	if conn.db != nil {
		rows, err := conn.db.Query(sqlText)
		if err != nil {
			vm.adodbConnectionRaiseProviderError(conn, "ADODB.Recordset.Open", err.Error(), "")
			return
		}
		rs.sqlRows = rows
		vm.adodbRecordsetLoadCurrentSQLResultSet(rs)
		return
	}

	if conn.oleConnection != nil {
		unknown, createErr := oleutil.CreateObject("ADODB.Recordset")
		if createErr != nil {
			vm.adodbConnectionRaiseProviderError(conn, "ADODB.Recordset.Open", "OLE: "+createErr.Error(), "")
			return
		}
		disp, dispatchErr := unknown.QueryInterface(ole.IID_IDispatch)
		unknown.Release()
		if dispatchErr != nil {
			vm.adodbConnectionRaiseProviderError(conn, "ADODB.Recordset.Open", "OLE: "+dispatchErr.Error(), "")
			return
		}

		cursorLocation, cursorType, lockType := vm.adodbResolveOLERecordsetOpenOptions(rs, conn, args)
		if _, putErr := oleutil.PutProperty(disp, "CursorLocation", cursorLocation); putErr != nil {
			disp.Release()
			vm.adodbConnectionRaiseProviderError(conn, "ADODB.Recordset.Open", "OLE: "+putErr.Error(), "")
			return
		}

		openArgs := make([]interface{}, 0, 5)
		openArgs = append(openArgs, sqlText, conn.oleConnection, cursorType, lockType)
		if len(args) >= 5 && args[4].Type != VTEmpty {
			openArgs = append(openArgs, vm.adodbOLEVariantArg(args[4]))
		}

		openRes, openErr := oleutil.CallMethod(disp, "Open", openArgs...)
		if openErr != nil {
			disp.Release()
			vm.adodbConnectionRaiseProviderError(conn, "ADODB.Recordset.Open", "OLE: "+openErr.Error(), "")
			return
		}
		if openRes != nil {
			openRes.Clear()
		}

		rs.oleRecordset = disp
		rs.cursorLocation = int(cursorLocation)
		rs.cursorType = int(cursorType)
		rs.lockType = int(lockType)
		rs.state = adStateOpen
		// Materialise rows immediately, then release the COM cursor so later VM work
		// stays in the in-memory cache and never depends on apartment-bound MoveNext calls.
		vm.adodbPopulateRecordsetFromOLE(rs)
		_, _ = oleutil.CallMethod(disp, "Close")
		disp.Release()
		rs.oleRecordset = nil
		rs.status = adRecOK
		rs.editMode = adEditNone
	}
}

// adodbResolveOLERecordsetOpenOptions computes the cursor options used for one OLE Recordset.Open call.
func (vm *VM) adodbResolveOLERecordsetOpenOptions(rs *adodbRecordset, conn *adodbConnection, args []Value) (int32, int32, int32) {
	cursorLocation := adUseServer
	if conn != nil && conn.cursorLocation != 0 {
		cursorLocation = conn.cursorLocation
	}
	if rs != nil && rs.cursorLocation != 0 {
		cursorLocation = rs.cursorLocation
	}

	cursorType := adOpenForwardOnly
	if rs != nil && rs.cursorType != 0 {
		cursorType = rs.cursorType
	}
	if len(args) >= 3 && args[2].Type != VTEmpty {
		cursorType = vm.asInt(args[2])
	}
	if len(args) < 3 && rs != nil && rs.activeConnection != 0 && (cursorType == adOpenKeyset || cursorType == adOpenDynamic) {
		// Access is far more reliable with a static cursor when VBScript sets
		// ActiveConnection and then calls Recordset.Open(sql) without explicit
		// cursor arguments. This preserves RecordCount/AbsolutePosition support
		// without the provider stalling on keyset/dynamic JOIN queries.
		cursorType = adOpenStatic
	}

	lockType := adLockReadOnly
	if rs != nil && rs.lockType != 0 {
		lockType = rs.lockType
	}
	if len(args) >= 4 && args[3].Type != VTEmpty {
		lockType = vm.asInt(args[3])
	}

	if cursorLocation == adUseClient && (cursorType == adOpenKeyset || cursorType == adOpenDynamic) {
		cursorType = adOpenStatic
	}

	return int32(cursorLocation), int32(cursorType), int32(lockType)
}

// adodbRecordsetRebuildColumnIndex rebuilds the lowercase column-name to ordinal map.
func (vm *VM) adodbRecordsetRebuildColumnIndex(rs *adodbRecordset) {
	if rs == nil {
		return
	}
	if rs.columnIndexByLower == nil {
		rs.columnIndexByLower = make(map[string]int, len(rs.columns))
	} else {
		clear(rs.columnIndexByLower)
	}
	for idx := 0; idx < len(rs.columns); idx++ {
		key := strings.ToLower(strings.TrimSpace(rs.columns[idx]))
		if key == "" {
			continue
		}
		if _, exists := rs.columnIndexByLower[key]; !exists {
			rs.columnIndexByLower[key] = idx
		}
	}
}

// adodbRecordsetClearPendingUpdateFields resets tracked changed columns for one recordset edit cycle.
func (vm *VM) adodbRecordsetClearPendingUpdateFields(rs *adodbRecordset) {
	if rs == nil {
		return
	}
	if rs.pendingUpdateFields == nil {
		rs.pendingUpdateFields = make(map[string]struct{})
		return
	}
	clear(rs.pendingUpdateFields)
}

// adodbRecordsetMarkPendingUpdateField tracks one changed field for Update writeback.
func (vm *VM) adodbRecordsetMarkPendingUpdateField(rs *adodbRecordset, fieldName string) {
	if rs == nil {
		return
	}
	key := strings.ToLower(strings.TrimSpace(fieldName))
	if key == "" {
		return
	}
	if rs.pendingUpdateFields == nil {
		rs.pendingUpdateFields = make(map[string]struct{})
	}
	rs.pendingUpdateFields[key] = struct{}{}
}

// adodbRecordsetResolveColumnKey normalizes a field selector to the row-map key format.
func (vm *VM) adodbRecordsetResolveColumnKey(rs *adodbRecordset, selector Value) string {
	if selector.Type == VTInteger || selector.Type == VTDouble {
		idx := vm.asInt(selector)
		if idx >= 0 && idx < len(rs.columns) {
			return strings.ToLower(strings.TrimSpace(rs.columns[idx]))
		}
		return ""
	}
	name := strings.ToLower(strings.TrimSpace(selector.String()))
	if name == "" {
		return ""
	}
	if rs != nil && rs.columnIndexByLower != nil {
		if _, exists := rs.columnIndexByLower[name]; exists {
			return name
		}
	}
	return name
}

// adodbOLEVariantArg converts one VM value to the compact OLE argument form accepted by oleutil.
func (vm *VM) adodbOLEVariantArg(val Value) interface{} {
	switch val.Type {
	case VTEmpty, VTNull:
		return nil
	case VTBool:
		return val.Num != 0
	case VTInteger:
		return int32(val.Num)
	case VTDouble:
		return val.Flt
	case VTString:
		return val.Str
	default:
		return val.String()
	}
}

// adodbRecordsetLoadCurrentSQLResultSet materializes the current sql.Rows result set into the Recordset cache.
func (vm *VM) adodbRecordsetLoadCurrentSQLResultSet(rs *adodbRecordset) {
	if rs == nil || rs.sqlRows == nil {
		rs.fieldChunkOffset = make(map[string]int)
		vm.adodbRecordsetClearPendingUpdateFields(rs)
		rs.columns = nil
		vm.adodbRecordsetRebuildColumnIndex(rs)
		rs.data = nil
		rs.recordCount = 0
		rs.currentRow = -1
		rs.state = adStateOpen
		rs.status = adRecOK
		rs.editMode = adEditNone
		rs.eof = true
		rs.bof = true
		return
	}

	rows := rs.sqlRows
	rs.fieldChunkOffset = make(map[string]int)
	vm.adodbRecordsetClearPendingUpdateFields(rs)
	cols, err := rows.Columns()
	if err != nil {
		cols = nil
	}
	if len(cols) == 0 {
		rs.columns = nil
	} else {
		rs.columns = cols
	}
	vm.adodbRecordsetRebuildColumnIndex(rs)
	rs.data = make([]map[string]Value, 0)

	for rows.Next() {
		if rs.maxRecords > 0 && len(rs.data) >= rs.maxRecords {
			break
		}
		rowValues := make([]interface{}, len(cols))
		rowPointers := make([]interface{}, len(cols))
		for i := range rowValues {
			rowPointers[i] = &rowValues[i]
		}
		if err := rows.Scan(rowPointers...); err != nil {
			continue
		}

		rowMap := make(map[string]Value, len(cols))
		for i, col := range cols {
			rowMap[strings.ToLower(col)] = vm.adodbValueToVMValue(rowValues[i])
		}
		rs.data = append(rs.data, rowMap)
	}

	rs.recordCount = len(rs.data)
	rs.state = adStateOpen
	rs.status = adRecOK
	rs.editMode = adEditNone
	rs.bookmark = 0
	if rs.recordCount > 0 {
		rs.currentRow = 0
		rs.eof = false
		rs.bof = false
		return
	}
	rs.currentRow = -1
	rs.eof = true
	rs.bof = true
}

// adodbRecordsetNextRecordset advances to the next SQL result set and returns a new Recordset object.
func (vm *VM) adodbRecordsetNextRecordset(rs *adodbRecordset) Value {
	if rs == nil || rs.sqlRows == nil {
		return Value{Type: VTObject, Num: 0}
	}
	if !rs.sqlRows.NextResultSet() {
		_ = rs.sqlRows.Close()
		rs.sqlRows = nil
		return Value{Type: VTObject, Num: 0}
	}

	nextVal := vm.newADODBRecordset()
	next := vm.adodbRecordsetItems[nextVal.Num]
	if next == nil {
		return Value{Type: VTObject, Num: 0}
	}

	next.activeCommand = rs.activeCommand
	next.activeConnection = rs.activeConnection
	next.cursorType = rs.cursorType
	next.lockType = rs.lockType
	next.cursorLocation = rs.cursorLocation
	next.source = rs.source
	next.pageSize = rs.pageSize
	next.cacheSize = rs.cacheSize
	next.maxRecords = rs.maxRecords
	next.sqlRows = rs.sqlRows
	vm.adodbRecordsetLoadCurrentSQLResultSet(next)
	return nextVal
}

func (vm *VM) adodbRecordsetClose(rs *adodbRecordset) {
	if rs.sqlRows != nil {
		_ = rs.sqlRows.Close()
		rs.sqlRows = nil
	}
	rs.state = adStateClosed
	rs.data = nil
	rs.columns = nil
	vm.adodbRecordsetRebuildColumnIndex(rs)
	rs.currentRow = -1
	rs.eof = true
	rs.bof = true
	rs.editMode = adEditNone
	rs.status = adRecOK
	rs.fieldChunkOffset = make(map[string]int)
	vm.adodbRecordsetClearPendingUpdateFields(rs)
	if rs.oleRecordset != nil {
		// Close the OLE Recordset before releasing to free server-side resources.
		_, _ = oleutil.CallMethod(rs.oleRecordset, "Close")
		rs.oleRecordset.Release()
		rs.oleRecordset = nil
	}
}

func (vm *VM) adodbRecordsetMoveNext(rs *adodbRecordset) {
	if rs.state != adStateOpen {
		return
	}
	if vm.adodbRecordsetIsLiveOLE(rs) {
		_, _ = oleutil.CallMethod(rs.oleRecordset, "MoveNext")
		vm.adodbOLERefreshCursorFlags(rs)
		if rs.eof {
			rs.currentRow = rs.recordCount
		} else {
			rs.currentRow++
			rs.bookmark = rs.currentRow + 1
		}
		return
	}
	rs.currentRow++
	if rs.currentRow >= rs.recordCount {
		rs.currentRow = rs.recordCount
		rs.eof = true
		return
	}
	rs.bookmark = rs.currentRow + 1
	rs.bof = false
}

func (vm *VM) adodbRecordsetMovePrevious(rs *adodbRecordset) {
	if rs.state != adStateOpen {
		return
	}
	if vm.adodbRecordsetIsLiveOLE(rs) {
		_, _ = oleutil.CallMethod(rs.oleRecordset, "MovePrevious")
		vm.adodbOLERefreshCursorFlags(rs)
		if rs.bof {
			rs.currentRow = -1
		} else if rs.currentRow > 0 {
			rs.currentRow--
			rs.bookmark = rs.currentRow + 1
		}
		return
	}
	rs.currentRow--
	if rs.currentRow < 0 {
		rs.currentRow = -1
		rs.bof = true
		return
	}
	rs.bookmark = rs.currentRow + 1
	rs.eof = false
}

func (vm *VM) adodbRecordsetMoveFirst(rs *adodbRecordset) {
	if rs.state != adStateOpen || rs.recordCount == 0 {
		return
	}
	if vm.adodbRecordsetIsLiveOLE(rs) {
		_, _ = oleutil.CallMethod(rs.oleRecordset, "MoveFirst")
		vm.adodbOLERefreshCursorFlags(rs)
		if !rs.eof {
			rs.currentRow = 0
			rs.bookmark = 1
		}
		return
	}
	rs.currentRow = 0
	rs.bookmark = 1
	rs.eof = false
	rs.bof = false
}

func (vm *VM) adodbRecordsetMoveLast(rs *adodbRecordset) {
	if rs.state != adStateOpen || rs.recordCount == 0 {
		return
	}
	if vm.adodbRecordsetIsLiveOLE(rs) {
		_, _ = oleutil.CallMethod(rs.oleRecordset, "MoveLast")
		vm.adodbOLERefreshCursorFlags(rs)
		if rs.recordCount > 0 && !rs.bof {
			rs.currentRow = rs.recordCount - 1
			rs.bookmark = rs.recordCount
		}
		return
	}
	rs.currentRow = rs.recordCount - 1
	rs.bookmark = rs.currentRow + 1
	rs.eof = false
	rs.bof = false
}

// adodbRecordsetMove moves by an offset relative to current row.
func (vm *VM) adodbRecordsetMove(rs *adodbRecordset, args []Value) {
	if rs.state != adStateOpen || len(args) == 0 {
		return
	}
	offset := vm.asInt(args[0])
	if vm.adodbRecordsetIsLiveOLE(rs) {
		_, _ = oleutil.CallMethod(rs.oleRecordset, "Move", int32(offset))
		vm.adodbOLERefreshCursorFlags(rs)
		if rs.bof {
			rs.currentRow = -1
		} else if rs.eof {
			rs.currentRow = rs.recordCount
		} else {
			rs.currentRow += offset
			if rs.currentRow < 0 {
				rs.currentRow = 0
			}
			rs.bookmark = rs.currentRow + 1
		}
		return
	}
	target := rs.currentRow + offset
	if target < 0 {
		rs.currentRow = -1
		rs.bof = true
		rs.eof = false
		return
	}
	if target >= rs.recordCount {
		rs.currentRow = rs.recordCount
		rs.eof = true
		rs.bof = false
		return
	}
	rs.currentRow = target
	rs.bookmark = rs.currentRow + 1
	rs.eof = false
	rs.bof = false
}

// adodbRecordsetFind performs a compact compatibility search for "field = value" criteria.
func (vm *VM) adodbRecordsetFind(rs *adodbRecordset, args []Value) {
	if rs.state != adStateOpen || len(args) == 0 || rs.recordCount == 0 {
		return
	}
	field, expected, ok := adodbParseFindCriteria(vm.valueToString(args[0]))
	if !ok {
		vm.raise(vbscript.InternalError, NewAxonASPError(ErrInvalidProcedureCallOrArg, nil, "ADODB.Recordset.Find criteria is invalid", "axonvm/lib_adodb.go", 0).Error())
		return
	}
	start := rs.currentRow + 1
	if start < 0 {
		start = 0
	}
	for i := start; i < len(rs.data); i++ {
		row := rs.data[i]
		if row == nil {
			continue
		}
		if val, exists := row[field]; exists && strings.EqualFold(strings.TrimSpace(val.String()), expected) {
			rs.currentRow = i
			rs.bookmark = i + 1
			rs.eof = false
			rs.bof = false
			return
		}
	}
	rs.currentRow = rs.recordCount
	rs.eof = true
}

// adodbRecordsetClone creates an in-memory clone preserving schema and current view state.
func (vm *VM) adodbRecordsetClone(rs *adodbRecordset) Value {
	cloneVal := vm.newADODBRecordset()
	clone := vm.adodbRecordsetItems[cloneVal.Num]
	if clone == nil {
		return Value{Type: VTEmpty}
	}
	clone.eof = rs.eof
	clone.bof = rs.bof
	clone.state = rs.state
	clone.currentRow = rs.currentRow
	clone.bookmark = rs.bookmark
	clone.recordCount = rs.recordCount
	clone.pageSize = rs.pageSize
	clone.absolutePage = rs.absolutePage
	clone.cacheSize = rs.cacheSize
	clone.filter = rs.filter
	clone.sort = rs.sort
	clone.dataMember = rs.dataMember
	clone.editMode = rs.editMode
	clone.index = rs.index
	clone.marshalOptions = rs.marshalOptions
	clone.maxRecords = rs.maxRecords
	clone.status = rs.status
	clone.activeCommand = rs.activeCommand
	clone.activeConnection = rs.activeConnection
	clone.cursorType = rs.cursorType
	clone.lockType = rs.lockType
	clone.cursorLocation = rs.cursorLocation
	clone.source = rs.source
	clone.columns = append(clone.columns[:0], rs.columns...)
	vm.adodbRecordsetRebuildColumnIndex(clone)
	clone.columnTypes = append(clone.columnTypes[:0], rs.columnTypes...)
	for k, v := range rs.columnTypeByName {
		clone.columnTypeByName[k] = v
	}
	for k, v := range rs.columnSizeByName {
		clone.columnSizeByName[k] = v
	}
	for k, v := range rs.columnAttrByName {
		clone.columnAttrByName[k] = v
	}
	for k, v := range rs.columnScaleByName {
		clone.columnScaleByName[k] = v
	}
	clone.data = make([]map[string]Value, len(rs.data))
	for i := 0; i < len(rs.data); i++ {
		if rs.data[i] == nil {
			continue
		}
		rowCopy := make(map[string]Value, len(rs.data[i]))
		for k, v := range rs.data[i] {
			rowCopy[k] = v
		}
		clone.data[i] = rowCopy
	}
	return cloneVal
}

// adodbRecordsetApplyFilterSort applies compact in-memory filter/sort compatibility behavior.
func (vm *VM) adodbRecordsetApplyFilterSort(rs *adodbRecordset) {
	if rs == nil {
		return
	}
	if rs.filter == "" && rs.sort == "" {
		rs.recordCount = len(rs.data)
		if rs.recordCount == 0 {
			rs.currentRow = -1
			rs.eof = true
			rs.bof = true
		} else if rs.currentRow < 0 || rs.currentRow >= rs.recordCount {
			rs.currentRow = 0
			rs.eof = false
			rs.bof = false
		}
		return
	}

	filtered := make([]map[string]Value, 0, len(rs.data))
	field := ""
	expected := ""
	hasFilter := false
	if strings.TrimSpace(rs.filter) != "" {
		var ok bool
		field, expected, ok = adodbParseFindCriteria(rs.filter)
		hasFilter = ok
	}

	for i := 0; i < len(rs.data); i++ {
		row := rs.data[i]
		if row == nil {
			continue
		}
		if hasFilter {
			val, exists := row[field]
			if !exists || !strings.EqualFold(strings.TrimSpace(val.String()), expected) {
				continue
			}
		}
		filtered = append(filtered, row)
	}
	if strings.TrimSpace(rs.sort) != "" {
		sortField, descending := adodbParseSortClause(rs.sort)
		if sortField != "" {
			sort.SliceStable(filtered, func(i, j int) bool {
				left := ""
				right := ""
				if filtered[i] != nil {
					left = strings.TrimSpace(filtered[i][sortField].String())
				}
				if filtered[j] != nil {
					right = strings.TrimSpace(filtered[j][sortField].String())
				}
				if descending {
					return left > right
				}
				return left < right
			})
		}
	}
	rs.data = filtered
	rs.recordCount = len(filtered)
	if rs.recordCount == 0 {
		rs.currentRow = -1
		rs.eof = true
		rs.bof = true
		return
	}
	rs.currentRow = 0
	rs.bookmark = 1
	rs.eof = false
	rs.bof = false
}

// adodbParseFindCriteria parses "field = value" style clauses used by Find and Filter.
func adodbParseFindCriteria(criteria string) (string, string, bool) {
	trimmed := strings.TrimSpace(criteria)
	if trimmed == "" {
		return "", "", false
	}
	idx := strings.Index(trimmed, "=")
	if idx <= 0 {
		return "", "", false
	}
	field := strings.ToLower(strings.TrimSpace(trimmed[:idx]))
	value := strings.TrimSpace(trimmed[idx+1:])
	value = strings.Trim(value, `"`)
	value = strings.Trim(value, `'`)
	if field == "" {
		return "", "", false
	}
	return field, value, true
}

// adodbParseSortClause parses a compact "field [ASC|DESC]" sort expression.
func adodbParseSortClause(clause string) (string, bool) {
	trimmed := strings.TrimSpace(clause)
	if trimmed == "" {
		return "", false
	}
	parts := strings.Fields(trimmed)
	if len(parts) == 0 {
		return "", false
	}
	field := strings.ToLower(strings.TrimSpace(parts[0]))
	desc := false
	if len(parts) > 1 {
		desc = strings.EqualFold(parts[1], "DESC")
	}
	return field, desc
}

func (vm *VM) adodbRecordsetAddNew(rs *adodbRecordset, args []Value) {
	vm.adodbRecordsetClearPendingUpdateFields(rs)
	if vm.adodbRecordsetIsLiveOLE(rs) {
		_, err := oleutil.CallMethod(rs.oleRecordset, "AddNew")
		if err != nil {
			if conn := vm.adodbRecordsetConnection(rs); conn != nil {
				vm.adodbConnectionRaiseProviderError(conn, "ADODB.Recordset.AddNew", "OLE: "+err.Error(), "")
				return
			}
			vm.raise(vbscript.TypeMismatch, "ADODB.Recordset.AddNew: OLE: "+err.Error())
			return
		}
		rs.editMode = adEditAdd
		rs.status = adRecOK
		vm.adodbOLERefreshPositionFlags(rs)
		return
	}

	// Simple in-memory AddNew for now
	newRow := make(map[string]Value)
	for _, col := range rs.columns {
		newRow[strings.ToLower(col)] = Value{Type: VTEmpty}
	}
	rs.data = append(rs.data, newRow)
	rs.recordCount = len(rs.data)
	rs.currentRow = rs.recordCount - 1
	rs.eof = false
	rs.bof = false
	rs.editMode = adEditAdd
	rs.status = adRecOK
}

// adodbRecordsetAppendField appends one field definition to the in-memory recordset schema.
func (vm *VM) adodbRecordsetAppendField(rs *adodbRecordset, name string, fieldType int, definedSize int, attributes int) {
	trimmedName := strings.TrimSpace(name)
	if trimmedName == "" {
		return
	}
	lowerName := strings.ToLower(trimmedName)
	if rs.columnIndexByLower == nil {
		vm.adodbRecordsetRebuildColumnIndex(rs)
	}
	if _, exists := rs.columnIndexByLower[lowerName]; exists {
		rs.columnTypeByName[lowerName] = fieldType
		rs.columnSizeByName[lowerName] = definedSize
		rs.columnAttrByName[lowerName] = attributes
		return
	}

	rs.columns = append(rs.columns, trimmedName)
	rs.columnTypes = append(rs.columnTypes, fmt.Sprintf("%d", fieldType))
	rs.columnTypeByName[lowerName] = fieldType
	rs.columnSizeByName[lowerName] = definedSize
	rs.columnAttrByName[lowerName] = attributes
	vm.adodbRecordsetRebuildColumnIndex(rs)

	for i := 0; i < len(rs.data); i++ {
		if rs.data[i] == nil {
			rs.data[i] = make(map[string]Value)
		}
		if _, exists := rs.data[i][lowerName]; !exists {
			rs.data[i][lowerName] = Value{Type: VTEmpty}
		}
	}
}

func (vm *VM) adodbRecordsetUpdate(rs *adodbRecordset, args []Value) {
	if rs == nil {
		return
	}
	if rs.state != adStateOpen {
		vm.raise(vbscript.CannotPerformTheRequestedOperation, "Operation is not allowed when the object is closed")
		return
	}
	if vm.adodbRecordsetIsLiveOLE(rs) {
		_, err := oleutil.CallMethod(rs.oleRecordset, "Update")
		if err != nil {
			if conn := vm.adodbRecordsetConnection(rs); conn != nil {
				vm.adodbConnectionRaiseProviderError(conn, "ADODB.Recordset.Update", "OLE: "+err.Error(), "")
				return
			}
			vm.raise(vbscript.TypeMismatch, "ADODB.Recordset.Update: OLE: "+err.Error())
			return
		}
		rs.editMode = adEditNone
		rs.status = adRecOK
		vm.adodbRecordsetClearPendingUpdateFields(rs)
		vm.adodbOLERefreshPositionFlags(rs)
		return
	}
	if rs.editMode == adEditAdd {
		if !vm.adodbPersistRecordsetInsert(rs) {
			return
		}
	} else {
		if !vm.adodbPersistRecordsetUpdate(rs) {
			return
		}
	}
	rs.editMode = adEditNone
	rs.status = adRecOK
	vm.adodbRecordsetClearPendingUpdateFields(rs)
}

func (vm *VM) adodbRecordsetDelete(rs *adodbRecordset) {
	if rs.state != adStateOpen || rs.currentRow < 0 || rs.currentRow >= rs.recordCount {
		return
	}
	rs.data = append(rs.data[:rs.currentRow], rs.data[rs.currentRow+1:]...)
	rs.recordCount = len(rs.data)
	if rs.recordCount == 0 {
		rs.currentRow = -1
		rs.eof = true
		rs.bof = true
	} else if rs.currentRow >= rs.recordCount {
		rs.currentRow = rs.recordCount - 1
	}
	rs.editMode = adEditNone
	rs.status = adRecOK
}

func (vm *VM) adodbRecordsetGetRows(rs *adodbRecordset, args []Value) Value {
	if rs.state != adStateOpen || rs.recordCount == 0 {
		return Value{Type: VTEmpty}
	}
	// Return a 2D array [cols][rows] as nested arrays
	rows := rs.recordCount
	cols := len(rs.columns)

	colArray := NewVBArray(0, cols)
	for c := 0; c < cols; c++ {
		rowArray := NewVBArray(0, rows)
		for r := 0; r < rows; r++ {
			rowArray.Values[r] = rs.data[r][strings.ToLower(rs.columns[c])]
		}
		colArray.Values[c] = Value{Type: VTArray, Arr: rowArray}
	}

	return Value{Type: VTArray, Arr: colArray}
}

func (vm *VM) adodbRecordsetGetString(rs *adodbRecordset, args []Value) Value {
	if rs.state != adStateOpen || rs.recordCount == 0 {
		return NewString("")
	}
	var sb strings.Builder
	for r := 0; r < rs.recordCount; r++ {
		row := rs.data[r]
		for c := 0; c < len(rs.columns); c++ {
			sb.WriteString(row[strings.ToLower(rs.columns[c])].String())
			if c < len(rs.columns)-1 {
				sb.WriteString("\t")
			}
		}
		sb.WriteString("\n")
	}
	return NewString(sb.String())
}

// --- Command Implementation ---

func (vm *VM) dispatchADODBCommandMethod(cmd *adodbCommand, member string, args []Value) Value {
	switch {
	case strings.EqualFold(member, "Execute"):
		return vm.adodbCommandExecute(cmd, args)
	case strings.EqualFold(member, "CreateParameter"):
		return vm.adodbCreateParameter(cmd, args)
	case strings.EqualFold(member, "Cancel"):
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Parameters.Refresh"):
		return Value{Type: VTEmpty}
	}
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchADODBCommandPropertyGet(cmd *adodbCommand, member string) Value {
	switch {
	case strings.EqualFold(member, "ActiveConnection"):
		return Value{Type: VTNativeObject, Num: cmd.activeConnection}
	case strings.EqualFold(member, "CommandText"):
		return NewString(cmd.commandText)
	case strings.EqualFold(member, "CommandType"):
		return NewInteger(int64(cmd.commandType))
	case strings.EqualFold(member, "CommandTimeout"):
		return NewInteger(int64(cmd.commandTimeout))
	case strings.EqualFold(member, "Prepared"):
		return NewBool(cmd.prepared)
	case strings.EqualFold(member, "Parameters"):
		return vm.newADODBParametersCollection(cmd)
	}
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchADODBCommandPropertySet(cmd *adodbCommand, member string, val Value) bool {
	switch {
	case strings.EqualFold(member, "ActiveConnection"):
		if val.Type == VTNativeObject {
			cmd.activeConnection = val.Num
		}
		return true
	case strings.EqualFold(member, "CommandText"):
		cmd.commandText = val.String()
		return true
	case strings.EqualFold(member, "CommandType"):
		cmd.commandType = vm.asInt(val)
		return true
	case strings.EqualFold(member, "CommandTimeout"):
		cmd.commandTimeout = vm.asInt(val)
		return true
	case strings.EqualFold(member, "Prepared"):
		cmd.prepared = vm.asBool(val)
		return true
	}
	return false
}

func (vm *VM) adodbCommandExecute(cmd *adodbCommand, args []Value) Value {
	connID := cmd.activeConnection
	if connID == 0 {
		vm.raise(vbscript.InternalError,
			"ADODB.Command.Execute: no active connection")
		return Value{Type: VTEmpty}
	}
	conn := vm.adodbConnectionItems[connID]
	if conn == nil {
		vm.raise(vbscript.InternalError,
			"ADODB.Command.Execute: invalid connection")
		return Value{Type: VTEmpty}
	}

	params := make([]Value, len(cmd.parameters))
	for i, p := range cmd.parameters {
		params[i] = p.value
	}

	// For now, reuse connection execute
	return vm.adodbConnectionExecute(conn, []Value{NewString(cmd.commandText)})
}

func (vm *VM) adodbCreateParameter(cmd *adodbCommand, args []Value) Value {
	param := &adodbParameter{}
	if len(args) > 0 {
		param.name = args[0].String()
	}
	if len(args) > 1 {
		param.typ = vm.asInt(args[1])
	}
	if len(args) > 2 {
		param.direction = vm.asInt(args[2])
	}
	if len(args) > 3 {
		param.size = vm.asInt(args[3])
	}
	if len(args) > 4 {
		param.value = args[4]
	}

	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.adodbParameterItems[objID] = param
	return Value{Type: VTNativeObject, Num: objID}
}

// --- Collections & Proxies ---

const (
	nativeADODBErrorsCollection int64 = 30000 + iota
	nativeADODBFieldsCollection
	nativeADODBParametersCollection
)

func (vm *VM) newADODBErrorsCollection(conn *adodbConnection) Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.adodbErrorsCollectionItems[objID] = conn
	return Value{Type: VTNativeObject, Num: objID}
}

func (vm *VM) newADODBFieldsCollection(rs *adodbRecordset) Value {
	// We need a unique ID for this collection linked to the recordset
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.adodbFieldsCollectionItems[objID] = rs
	return Value{Type: VTNativeObject, Num: objID}
}

func (vm *VM) newADODBParametersCollection(cmd *adodbCommand) Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.adodbParametersCollectionItems[objID] = cmd
	return Value{Type: VTNativeObject, Num: objID}
}

func (vm *VM) dispatchADODBErrorsCollectionMethod(objID int64, member string, args []Value) (Value, bool) {
	conn, exists := vm.adodbErrorsCollectionItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}

	switch {
	case strings.EqualFold(member, "Count"):
		return NewInteger(int64(len(conn.errors))), true
	case member == "" || strings.EqualFold(member, "Item"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}, true
		}
		idx := vm.asInt(args[0])
		if idx < 0 || idx >= len(conn.errors) {
			return Value{Type: VTEmpty}, true
		}
		errObjID := vm.nextDynamicNativeID
		vm.nextDynamicNativeID++
		errCopy := conn.errors[idx]
		vm.adodbErrorItems[errObjID] = &errCopy
		return Value{Type: VTNativeObject, Num: errObjID}, true
	case strings.EqualFold(member, "Clear"):
		vm.adodbConnectionClearErrors(conn)
		return Value{Type: VTEmpty}, true
	}
	return Value{Type: VTEmpty}, false
}

func (vm *VM) dispatchADODBErrorsCollectionPropertyGet(objID int64, member string) (Value, bool) {
	conn, exists := vm.adodbErrorsCollectionItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}
	if strings.EqualFold(member, "Count") {
		return NewInteger(int64(len(conn.errors))), true
	}
	if strings.EqualFold(member, "Item") {
		return Value{Type: VTNativeObject, Num: objID}, true
	}
	return Value{Type: VTEmpty}, true
}

func (vm *VM) dispatchADODBErrorPropertyGet(objID int64, member string) (Value, bool) {
	errObj, exists := vm.adodbErrorItems[objID]
	if !exists || errObj == nil {
		return Value{Type: VTEmpty}, false
	}
	if strings.EqualFold(member, "Number") {
		return NewInteger(int64(errObj.number)), true
	}
	if strings.EqualFold(member, "Description") {
		return NewString(errObj.description), true
	}
	if strings.EqualFold(member, "Source") {
		return NewString(errObj.source), true
	}
	if strings.EqualFold(member, "SQLState") {
		return NewString(errObj.sqlState), true
	}
	return Value{Type: VTEmpty}, true
}

func (vm *VM) adodbConnectionClearErrors(conn *adodbConnection) {
	if conn == nil {
		return
	}
	conn.errors = conn.errors[:0]
}

func (vm *VM) adodbConnectionPushError(conn *adodbConnection, source string, number int, description string, sqlState string) {
	if conn == nil {
		return
	}
	conn.errors = append(conn.errors, adodbError{
		number:      number,
		description: description,
		source:      source,
		sqlState:    sqlState,
	})
}

func (vm *VM) adodbConnectionRaiseProviderError(conn *adodbConnection, source string, description string, sqlState string) {
	message := strings.TrimSpace(description)
	if message == "" {
		message = "Provider error"
	}
	number := vbscript.HRESULTFromVBScriptCode(vbscript.AutomationError)
	vm.adodbConnectionPushError(conn, source, number, message, sqlState)
	vm.raise(vbscript.AutomationError, source+": "+message)
}

func (vm *VM) dispatchADODBFieldsCollectionMethod(objID int64, member string, args []Value) (Value, bool) {
	rs, exists := vm.adodbFieldsCollectionItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}
	switch {
	case member == "" || strings.EqualFold(member, "Item"):
		if len(args) > 0 {
			if args[0].Type == VTInteger {
				idx := int(args[0].Num)
				if idx >= 0 && idx < len(rs.columns) {
					return vm.newADODBFieldProxy(rs, rs.columns[idx]), true
				}
				return Value{Type: VTEmpty}, true
			}
			name := args[0].String()
			return vm.newADODBFieldProxy(rs, name), true
		}
	case strings.EqualFold(member, "Append"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}, true
		}
		name := args[0].String()
		fieldType := 0
		definedSize := 0
		attributes := 0
		if len(args) > 1 && args[1].Type != VTEmpty {
			fieldType = vm.asInt(args[1])
		}
		if len(args) > 2 && args[2].Type != VTEmpty {
			definedSize = vm.asInt(args[2])
		}
		if len(args) > 3 && args[3].Type != VTEmpty {
			attributes = vm.asInt(args[3])
		}
		vm.adodbRecordsetAppendField(rs, name, fieldType, definedSize, attributes)
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "Count"):
		return NewInteger(int64(len(rs.columns))), true
	}
	return Value{Type: VTEmpty}, false
}

func (vm *VM) dispatchADODBParametersCollectionMethod(objID int64, member string, args []Value) (Value, bool) {
	cmd, exists := vm.adodbParametersCollectionItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}
	switch {
	case strings.EqualFold(member, "Append"):
		if len(args) > 0 && args[0].Type == VTNativeObject {
			if p, ok := vm.adodbParameterItems[args[0].Num]; ok {
				cmd.parameters = append(cmd.parameters, p)
			}
		}
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "Count"):
		return NewInteger(int64(len(cmd.parameters))), true
	}
	return Value{Type: VTEmpty}, false
}

// dispatchADODBFieldsCollectionPropertyGet handles property-get access on the
// ADODB.Fields collection object returned by Recordset.Fields.
// This is required because OpMemberGet uses dispatchMemberGet (property path),
// not dispatchNativeCall (method path). Without this handler, rsAccess.Fields.Count
// resolves to VTEmpty instead of the correct column count.
func (vm *VM) dispatchADODBFieldsCollectionPropertyGet(objID int64, member string) (Value, bool) {
	rs, exists := vm.adodbFieldsCollectionItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}
	switch {
	case strings.EqualFold(member, "Count"):
		return NewInteger(int64(len(rs.columns))), true
	case member == "" || strings.EqualFold(member, "Item"):
		// Default property: return the collection itself so chained access still works.
		return Value{Type: VTNativeObject, Num: objID}, true
	}
	return Value{Type: VTEmpty}, true
}

// dispatchADODBParametersCollectionPropertyGet handles property-get access on the
// ADODB.Parameters collection object returned by Command.Parameters.
func (vm *VM) dispatchADODBParametersCollectionPropertyGet(objID int64, member string) (Value, bool) {
	cmd, exists := vm.adodbParametersCollectionItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}
	if strings.EqualFold(member, "Count") {
		return NewInteger(int64(len(cmd.parameters))), true
	}
	if strings.EqualFold(member, "Item") {
		return Value{Type: VTNativeObject, Num: objID}, true
	}
	return Value{Type: VTEmpty}, true
}

func (vm *VM) newADODBFieldProxy(rs *adodbRecordset, name string) Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	lowerName := strings.ToLower(strings.TrimSpace(name))
	vm.adodbFieldItems[objID] = &adodbFieldProxy{rs: rs, name: name, cachedLowerName: lowerName}
	return Value{Type: VTNativeObject, Num: objID}
}

type adodbFieldProxy struct {
	rs              *adodbRecordset
	name            string
	cachedLowerName string
}

// adodbFieldChunkKey builds one stable key for chunk cursor state per field and row.
func (vm *VM) adodbFieldChunkKey(field *adodbFieldProxy) string {
	if field == nil || field.rs == nil {
		return ""
	}
	if field.rs.currentRow < 0 {
		return ""
	}
	return strconv.Itoa(field.rs.currentRow) + "|" + field.cachedLowerName
}

// adodbFieldChunkOffsetGet returns the current chunk cursor for one field on the active row.
func (vm *VM) adodbFieldChunkOffsetGet(field *adodbFieldProxy) int {
	if field == nil || field.rs == nil {
		return 0
	}
	if field.rs.fieldChunkOffset == nil {
		field.rs.fieldChunkOffset = make(map[string]int)
	}
	key := vm.adodbFieldChunkKey(field)
	if key == "" {
		return 0
	}
	offset, ok := field.rs.fieldChunkOffset[key]
	if !ok || offset < 0 {
		return 0
	}
	return offset
}

// adodbFieldChunkOffsetSet updates the chunk cursor for one field on the active row.
func (vm *VM) adodbFieldChunkOffsetSet(field *adodbFieldProxy, offset int) {
	if field == nil || field.rs == nil {
		return
	}
	if field.rs.fieldChunkOffset == nil {
		field.rs.fieldChunkOffset = make(map[string]int)
	}
	key := vm.adodbFieldChunkKey(field)
	if key == "" {
		return
	}
	if offset < 0 {
		offset = 0
	}
	field.rs.fieldChunkOffset[key] = offset
}

func (vm *VM) dispatchADODBFieldMethod(objID int64, member string, args []Value) (Value, bool) {
	field, exists := vm.adodbFieldItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}
	switch {
	case strings.EqualFold(member, "AppendChunk"):
		if len(args) >= 1 {
			current, _ := vm.dispatchADODBFieldPropertyGet(objID, "Value")
			combined := current.String() + args[0].String()
			vm.dispatchADODBFieldPropertySet(objID, "Value", NewString(combined))
			field.rs.editMode = adEditInProgress
		}
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "GetChunk"):
		value, _ := vm.dispatchADODBFieldPropertyGet(objID, "Value")
		all := value.String()
		offset := vm.adodbFieldChunkOffsetGet(field)
		if offset > len(all) {
			offset = len(all)
		}
		want := len(all) - offset
		if len(args) >= 1 {
			requested := vm.asInt(args[0])
			if requested < 0 {
				requested = 0
			}
			if requested < want {
				want = requested
			}
		}
		if want <= 0 {
			return NewString(""), true
		}
		end := offset + want
		if end > len(all) {
			end = len(all)
		}
		vm.adodbFieldChunkOffsetSet(field, end)
		return NewString(all[offset:end]), true
	}
	if member == "" {
		if field.rs.state == adStateOpen && field.rs.currentRow >= 0 && field.rs.currentRow < len(field.rs.data) {
			row := field.rs.data[field.rs.currentRow]
			return row[field.cachedLowerName], true
		}
	}
	return Value{Type: VTEmpty}, false
}

func (vm *VM) dispatchADODBFieldPropertyGet(objID int64, member string) (Value, bool) {
	field, exists := vm.adodbFieldItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}
	switch {
	case strings.EqualFold(member, "Value") || strings.EqualFold(member, "__default__") || member == "":
		if vm.adodbRecordsetIsLiveOLE(field.rs) {
			if val, ok := vm.adodbOLEGetFieldValue(field.rs, field.name); ok {
				return val, true
			}
			return Value{Type: VTEmpty}, true
		}
		if field.rs.state == adStateOpen && field.rs.currentRow >= 0 && field.rs.currentRow < len(field.rs.data) {
			row := field.rs.data[field.rs.currentRow]
			return row[field.cachedLowerName], true
		}
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "Name"):
		return NewString(field.name), true
	case strings.EqualFold(member, "Type"):
		return NewInteger(int64(field.rs.columnTypeByName[field.cachedLowerName])), true
	case strings.EqualFold(member, "DefinedSize"):
		return NewInteger(int64(field.rs.columnSizeByName[field.cachedLowerName])), true
	case strings.EqualFold(member, "Attributes"):
		return NewInteger(int64(field.rs.columnAttrByName[field.cachedLowerName])), true
	case strings.EqualFold(member, "NumericScale"):
		return NewInteger(int64(field.rs.columnScaleByName[field.cachedLowerName])), true
	case strings.EqualFold(member, "ActualSize"):
		value, _ := vm.dispatchADODBFieldPropertyGet(objID, "Value")
		return NewInteger(int64(len(value.String()))), true
	case strings.EqualFold(member, "DataFormat"):
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "OriginalValue"):
		value, _ := vm.dispatchADODBFieldPropertyGet(objID, "Value")
		return value, true
	case strings.EqualFold(member, "Precision"):
		precision := field.rs.columnSizeByName[field.cachedLowerName]
		if precision <= 0 {
			value, _ := vm.dispatchADODBFieldPropertyGet(objID, "Value")
			precision = len(value.String())
		}
		return NewInteger(int64(precision)), true
	case strings.EqualFold(member, "Status"):
		return NewInteger(int64(field.rs.status)), true
	case strings.EqualFold(member, "UnderlyingValue"):
		value, _ := vm.dispatchADODBFieldPropertyGet(objID, "Value")
		return value, true
	}
	return Value{Type: VTEmpty}, false
}

// adodbRecordsetIsLiveOLE reports whether this recordset should be served from the live OLE cursor.
// In this mode rows are not fully materialized in Go memory.
func (vm *VM) adodbRecordsetIsLiveOLE(rs *adodbRecordset) bool {
	return rs != nil && rs.oleRecordset != nil && rs.sqlRows == nil && rs.data == nil
}

// adodbInitOLERecordsetMetadata loads only schema and cursor flags from OLE, avoiding full row caching.
func (vm *VM) adodbInitOLERecordsetMetadata(rs *adodbRecordset) {
	if rs == nil || rs.oleRecordset == nil {
		return
	}
	rs.data = nil
	rs.fieldChunkOffset = make(map[string]int)

	fieldsRes, _ := oleutil.GetProperty(rs.oleRecordset, "Fields")
	if fieldsRes != nil {
		fields := fieldsRes.ToIDispatch()
		fieldsRes.Clear()
		if fields != nil {
			count := 0
			countRes, _ := oleutil.GetProperty(fields, "Count")
			if countRes != nil {
				count = vm.asInt(vm.adodbValueToVMValue(countRes.Value()))
				countRes.Clear()
			}
			rs.columns = make([]string, count)
			for i := 0; i < count; i++ {
				itemRes, _ := oleutil.GetProperty(fields, "Item", i)
				if itemRes == nil {
					continue
				}
				item := itemRes.ToIDispatch()
				itemRes.Clear()
				if item == nil {
					continue
				}
				nameRes, _ := oleutil.GetProperty(item, "Name")
				if nameRes != nil {
					name := strings.TrimSpace(nameRes.ToString())
					if name == "" {
						name = "field" + strconv.Itoa(i)
					}
					rs.columns[i] = name
					nameRes.Clear()
				}
				item.Release()
			}
			fields.Release()
		}
	}
	vm.adodbRecordsetRebuildColumnIndex(rs)

	countRes, _ := oleutil.GetProperty(rs.oleRecordset, "RecordCount")
	if countRes != nil {
		rs.recordCount = vm.asInt(vm.adodbValueToVMValue(countRes.Value()))
		countRes.Clear()
	}
	vm.adodbOLERefreshPositionFlags(rs)
}

// adodbOLERefreshCursorFlags synchronizes EOF/BOF flags from the live OLE cursor without forcing position reads.
func (vm *VM) adodbOLERefreshCursorFlags(rs *adodbRecordset) {
	if rs == nil || rs.oleRecordset == nil {
		return
	}
	eofRes, _ := oleutil.GetProperty(rs.oleRecordset, "EOF")
	if eofRes != nil {
		rs.eof = vm.asBool(vm.adodbValueToVMValue(eofRes.Value()))
		eofRes.Clear()
	}
	bofRes, _ := oleutil.GetProperty(rs.oleRecordset, "BOF")
	if bofRes != nil {
		rs.bof = vm.asBool(vm.adodbValueToVMValue(bofRes.Value()))
		bofRes.Clear()
	}
}

// adodbOLERefreshPositionFlags synchronizes EOF/BOF/current row from the live OLE cursor.
func (vm *VM) adodbOLERefreshPositionFlags(rs *adodbRecordset) {
	if rs == nil || rs.oleRecordset == nil {
		return
	}
	vm.adodbOLERefreshCursorFlags(rs)
	absRes, _ := oleutil.GetProperty(rs.oleRecordset, "AbsolutePosition")
	if absRes != nil {
		pos := vm.asInt(vm.adodbValueToVMValue(absRes.Value()))
		if pos > 0 {
			rs.currentRow = pos - 1
			rs.bookmark = pos
		} else if rs.eof && rs.recordCount >= 0 {
			rs.currentRow = rs.recordCount
		} else if rs.bof {
			rs.currentRow = -1
		}
		absRes.Clear()
	}
}

// adodbOLEGetFieldValue fetches one field value from the current row of a live OLE cursor.
func (vm *VM) adodbOLEGetFieldValue(rs *adodbRecordset, fieldName string) (Value, bool) {
	if rs == nil || rs.oleRecordset == nil {
		return Value{Type: VTEmpty}, false
	}
	fieldsRes, _ := oleutil.GetProperty(rs.oleRecordset, "Fields")
	if fieldsRes == nil {
		return Value{Type: VTEmpty}, false
	}
	fields := fieldsRes.ToIDispatch()
	if fields != nil {
		fields.AddRef()
	}
	fieldsRes.Clear()
	if fields == nil {
		return Value{Type: VTEmpty}, false
	}
	defer fields.Release()

	itemRes, _ := oleutil.GetProperty(fields, "Item", fieldName)
	if itemRes == nil {
		return Value{Type: VTEmpty}, false
	}
	item := itemRes.ToIDispatch()
	if item != nil {
		item.AddRef()
	}
	itemRes.Clear()
	if item == nil {
		return Value{Type: VTEmpty}, false
	}
	defer item.Release()

	valRes, _ := oleutil.GetProperty(item, "Value")
	if valRes == nil {
		return Value{Type: VTEmpty}, false
	}
	v := vm.adodbValueToVMValue(valRes.Value())
	valRes.Clear()
	return v, true
}

// adodbOLESetFieldValue writes one field value into the current row of a live OLE cursor.
func (vm *VM) adodbOLESetFieldValue(rs *adodbRecordset, fieldName string, val Value) bool {
	if rs == nil || rs.oleRecordset == nil {
		return false
	}
	fieldsRes, _ := oleutil.GetProperty(rs.oleRecordset, "Fields")
	if fieldsRes == nil {
		return false
	}
	fields := fieldsRes.ToIDispatch()
	if fields != nil {
		fields.AddRef()
	}
	fieldsRes.Clear()
	if fields == nil {
		return false
	}
	defer fields.Release()

	itemRes, _ := oleutil.GetProperty(fields, "Item", fieldName)
	if itemRes == nil {
		return false
	}
	item := itemRes.ToIDispatch()
	if item != nil {
		item.AddRef()
	}
	itemRes.Clear()
	if item == nil {
		return false
	}
	defer item.Release()

	setRes, setErr := oleutil.PutProperty(item, "Value", vm.adodbOLEVariantArg(val))
	if setRes != nil {
		setRes.Clear()
	}
	if setErr != nil {
		if conn := vm.adodbRecordsetConnection(rs); conn != nil {
			vm.adodbConnectionRaiseProviderError(conn, "ADODB.Recordset.Update", "OLE: "+setErr.Error(), "")
			return false
		}
		vm.raise(vbscript.TypeMismatch, "ADODB.Recordset.Update: OLE: "+setErr.Error())
		return false
	}

	return true
}

// dispatchADODBFieldPropertySet handles writable ADODB.Field properties.
func (vm *VM) dispatchADODBFieldPropertySet(objID int64, member string, val Value) bool {
	field, exists := vm.adodbFieldItems[objID]
	if !exists {
		return false
	}
	lowerName := field.cachedLowerName
	switch {
	case strings.EqualFold(member, "Value") || strings.EqualFold(member, "__default__") || member == "":
		if vm.adodbRecordsetIsLiveOLE(field.rs) {
			if vm.adodbOLESetFieldValue(field.rs, field.name, val) {
				if field.rs.editMode != adEditAdd {
					field.rs.editMode = adEditInProgress
				}
				vm.adodbRecordsetMarkPendingUpdateField(field.rs, field.name)
				vm.adodbFieldChunkOffsetSet(field, 0)
				return true
			}
			return false
		}
		if field.rs.state == adStateOpen && field.rs.currentRow >= 0 && field.rs.currentRow < len(field.rs.data) {
			row := field.rs.data[field.rs.currentRow]
			if row != nil {
				row[lowerName] = val
				if field.rs.editMode != adEditAdd {
					field.rs.editMode = adEditInProgress
				}
				vm.adodbRecordsetMarkPendingUpdateField(field.rs, lowerName)
				vm.adodbFieldChunkOffsetSet(field, 0)
			}
		}
		return true
	case strings.EqualFold(member, "NumericScale"):
		field.rs.columnScaleByName[lowerName] = vm.asInt(val)
		return true
	}
	return false
}

var adodbSimpleSelectSourceRE = regexp.MustCompile(`(?is)^\s*select\s+\*\s+from\s+(.+?)(?:\s+where\s+(.+?))?(?:\s+order\s+by\s+.+)?\s*$`)

// adodbPersistRecordsetUpdate writes the current row back to the backing table for simple single-table sources.
// For disconnected (client-side) recordsets that have no active connection, the edit is already in-memory
// and Update is a no-op — Classic ADO behaves identically in this case.
func (vm *VM) adodbPersistRecordsetUpdate(rs *adodbRecordset) bool {
	if rs.currentRow < 0 || rs.currentRow >= len(rs.data) {
		vm.raise(vbscript.InvalidProcedureCallOrArgument, "Either BOF or EOF is True, or the current record has been deleted")
		return false
	}
	conn := vm.adodbRecordsetConnection(rs)
	if conn == nil {
		// Disconnected (client-side) recordset — the row is already updated in rs.data; nothing to persist.
		return true
	}
	tableName, whereClause, keyField, ok := vm.adodbResolveUpdatableSource(rs.source)
	if !ok || strings.TrimSpace(whereClause) == "" || strings.EqualFold(strings.TrimSpace(whereClause), "1=2") {
		vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, "ADODB.Recordset.Update: source is not one simple updatable single-table query")
		return false
	}
	row := rs.data[rs.currentRow]
	assignments := make([]string, 0, len(rs.columns))
	usePending := len(rs.pendingUpdateFields) > 0
	for i := 0; i < len(rs.columns); i++ {
		column := rs.columns[i]
		lowerColumn := strings.ToLower(strings.TrimSpace(column))
		if usePending {
			if _, changed := rs.pendingUpdateFields[lowerColumn]; !changed {
				continue
			}
		}
		if keyField != "" && strings.EqualFold(column, keyField) {
			continue
		}
		literal, ok := vm.adodbSQLLiteral(conn, row[strings.ToLower(column)])
		if !ok {
			vm.raise(vbscript.TypeMismatch, "ADODB.Recordset.Update: unsupported field value type")
			return false
		}
		assignments = append(assignments, vm.adodbQuoteIdentifier(conn, column)+" = "+literal)
	}
	if len(assignments) == 0 {
		return true
	}
	sqlText := "UPDATE " + vm.adodbQuoteIdentifier(conn, tableName) + " SET " + strings.Join(assignments, ", ") + " WHERE " + whereClause
	if _, ok := vm.adodbExecWriteback(conn, sqlText, "ADODB.Recordset.Update", false); !ok {
		return false
	}
	return true
}

// adodbPersistRecordsetInsert inserts the current AddNew row back to the backing table for simple single-table sources.
// For disconnected (client-side) recordsets that have no active connection, the new row is already in-memory
// and Update/AddNew is a no-op from a persistence standpoint — Classic ADO behaves identically.
func (vm *VM) adodbPersistRecordsetInsert(rs *adodbRecordset) bool {
	if rs.currentRow < 0 || rs.currentRow >= len(rs.data) {
		vm.raise(vbscript.InvalidProcedureCallOrArgument, "Either BOF or EOF is True, or the current record has been deleted")
		return false
	}
	conn := vm.adodbRecordsetConnection(rs)
	if conn == nil {
		// Disconnected (client-side) recordset — the new row is already appended to rs.data; nothing to persist.
		return true
	}
	tableName, _, keyField, ok := vm.adodbResolveUpdatableSource(rs.source)
	if !ok {
		vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, "ADODB.Recordset.Update: source is not one simple updatable single-table query")
		return false
	}
	row := rs.data[rs.currentRow]
	columns := make([]string, 0, len(rs.columns))
	values := make([]string, 0, len(rs.columns))
	for i := 0; i < len(rs.columns); i++ {
		column := rs.columns[i]
		value := row[strings.ToLower(column)]
		if value.Type == VTEmpty {
			continue
		}
		if keyField != "" && strings.EqualFold(column, keyField) && (value.Type == VTEmpty || value.Type == VTNull || (value.Type == VTInteger && value.Num == 0) || (value.Type == VTString && strings.TrimSpace(value.Str) == "")) {
			continue
		}
		literal, ok := vm.adodbSQLLiteral(conn, value)
		if !ok {
			vm.raise(vbscript.TypeMismatch, "ADODB.Recordset.Update: unsupported field value type")
			return false
		}
		columns = append(columns, vm.adodbQuoteIdentifier(conn, column))
		values = append(values, literal)
	}
	if len(columns) == 0 {
		vm.raise(vbscript.TypeMismatch, "ADODB.Recordset.Update: no values available for insert")
		return false
	}
	sqlText := "INSERT INTO " + vm.adodbQuoteIdentifier(conn, tableName) + " (" + strings.Join(columns, ", ") + ") VALUES (" + strings.Join(values, ", ") + ")"
	wantIdentity := keyField != ""
	insertID, ok := vm.adodbExecWriteback(conn, sqlText, "ADODB.Recordset.Update", wantIdentity)
	if !ok {
		return false
	}
	if keyField != "" && insertID != 0 {
		row[strings.ToLower(keyField)] = NewInteger(insertID)
	}
	return true
}

// adodbRecordsetConnection resolves the active ADODB.Connection bound to one Recordset.
func (vm *VM) adodbRecordsetConnection(rs *adodbRecordset) *adodbConnection {
	if rs == nil || rs.activeConnection == 0 {
		return nil
	}
	return vm.adodbConnectionItems[rs.activeConnection]
}

// adodbResolveUpdatableSource extracts the base table and where clause from a simple SELECT * FROM table [WHERE ...] source.
func (vm *VM) adodbResolveUpdatableSource(source string) (string, string, string, bool) {
	trimmed := strings.TrimSpace(source)
	if trimmed == "" {
		return "", "", "", false
	}
	if vm.adodbLooksLikeSingleIdentifier(trimmed) {
		keyField := vm.adodbInferIdentityColumn([]string{trimmed})
		return trimmed, "", keyField, true
	}
	matches := adodbSimpleSelectSourceRE.FindStringSubmatch(trimmed)
	if len(matches) < 2 {
		return "", "", "", false
	}
	tableExpr := strings.TrimSpace(matches[1])
	if strings.ContainsAny(tableExpr, ",()") || strings.Contains(strings.ToLower(tableExpr), " join ") {
		return "", "", "", false
	}
	tableName := vm.adodbUnquoteIdentifier(tableExpr)
	if tableName == "" {
		return "", "", "", false
	}
	whereClause := ""
	if len(matches) > 2 {
		whereClause = strings.TrimSpace(matches[2])
	}
	keyField := vm.adodbInferKeyFieldFromWhere(whereClause)
	return tableName, whereClause, keyField, true
}

// adodbInferKeyFieldFromWhere extracts one equality column name from a simple WHERE clause.
func (vm *VM) adodbInferKeyFieldFromWhere(whereClause string) string {
	trimmed := strings.TrimSpace(whereClause)
	if trimmed == "" {
		return ""
	}
	parts := strings.Split(trimmed, "=")
	if len(parts) != 2 {
		return ""
	}
	left := strings.TrimSpace(parts[0])
	if strings.ContainsAny(strings.ToLower(left), " <>!()+-/*") {
		return ""
	}
	return vm.adodbUnquoteIdentifier(left)
}

// adodbInferIdentityColumn selects the most likely auto-increment key column name.
func (vm *VM) adodbInferIdentityColumn(columns []string) string {
	for i := 0; i < len(columns); i++ {
		if strings.EqualFold(columns[i], "id") || strings.EqualFold(columns[i], "iid") {
			return columns[i]
		}
	}
	return ""
}

// adodbUnquoteIdentifier removes one layer of common SQL identifier quoting and trailing table qualification.
func (vm *VM) adodbUnquoteIdentifier(identifier string) string {
	trimmed := strings.TrimSpace(identifier)
	if trimmed == "" {
		return ""
	}
	parts := strings.Split(trimmed, ".")
	trimmed = strings.TrimSpace(parts[len(parts)-1])
	if len(trimmed) >= 2 {
		first := trimmed[0]
		last := trimmed[len(trimmed)-1]
		if (first == '[' && last == ']') || (first == '`' && last == '`') || (first == '"' && last == '"') {
			trimmed = trimmed[1 : len(trimmed)-1]
		}
	}
	return strings.TrimSpace(trimmed)
}

// adodbSQLLiteral formats one VM value as one SQL literal compatible with the bound connection.
func (vm *VM) adodbSQLLiteral(conn *adodbConnection, val Value) (string, bool) {
	switch val.Type {
	case VTEmpty, VTNull:
		return "NULL", true
	case VTBool:
		if val.Num != 0 {
			return "1", true
		}
		return "0", true
	case VTInteger:
		return strconv.FormatInt(val.Num, 10), true
	case VTDouble:
		return strconv.FormatFloat(val.Flt, 'f', -1, 64), true
	case VTDate:
		normalized := vm.adodbBestEffortVMDate(val)
		if normalized.IsZero() {
			return "NULL", true
		}
		return vm.adodbDateLiteral(conn, normalized), true
	case VTString:
		if parsed, ok := vm.adodbParseDateLiteral(val.Str); ok {
			return vm.adodbDateLiteral(conn, parsed), true
		}
		escaped := strings.ReplaceAll(val.Str, "'", "''")
		return "'" + escaped + "'", true
	default:
		if val.Str != "" {
			escaped := strings.ReplaceAll(val.String(), "'", "''")
			return "'" + escaped + "'", true
		}
	}
	return "", false
}

// adodbBestEffortVMDate converts one VM VTDate payload to time.Time using a
// conservative fallback chain. The primary representation is Unix nanoseconds,
// but this helper also tolerates legacy millisecond/second payloads.
func (vm *VM) adodbBestEffortVMDate(val Value) time.Time {
	if val.Type != VTDate {
		return time.Time{}
	}
	if val.Num == NewDate(time.Time{}).Num {
		return time.Time{}
	}
	primary := time.Unix(0, val.Num).UTC()
	if primary.Year() >= 1800 && primary.Year() <= 9999 {
		return primary
	}
	if val.Num > -62135596800000 && val.Num < 253402300800000 {
		ms := time.UnixMilli(val.Num).UTC()
		if ms.Year() >= 1800 && ms.Year() <= 9999 {
			return ms
		}
	}
	if val.Num > -62135596800 && val.Num < 253402300800 {
		sec := time.Unix(val.Num, 0).UTC()
		if sec.Year() >= 1800 && sec.Year() <= 9999 {
			return sec
		}
	}
	return primary
}

// adodbDateLiteral formats one SQL date/time literal according to the active provider.
func (vm *VM) adodbDateLiteral(conn *adodbConnection, value time.Time) string {
	formatted := value.Format("2006-01-02 15:04:05")
	if conn != nil && (strings.Contains(strings.ToLower(conn.provider), "ace") || strings.Contains(strings.ToLower(conn.provider), "jet")) {
		// Access/Jet date literals are reliably parsed in U.S. month/day order.
		return "#" + value.Format("01/02/2006 15:04:05") + "#"
	}
	return "'" + formatted + "'"
}

// adodbParseDateLiteral parses the common VBScript/ADO date string formats used by aspLite samples.
func (vm *VM) adodbParseDateLiteral(text string) (time.Time, bool) {
	trimmed := strings.TrimSpace(text)
	if trimmed == "" {
		return time.Time{}, false
	}
	layouts := []string{
		"2006-01-02 15:04:05",
		"2006-01-02T15:04:05",
		time.RFC3339,
		"2006-01-02",
		"02/01/2006 15:04",
		"02/01/2006 15:04:05",
		"02/01/2006",
		"01/02/2006 15:04",
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

// adodbExecWriteback executes one non-query writeback statement and optionally returns the inserted identity.
func (vm *VM) adodbExecWriteback(conn *adodbConnection, sqlText string, source string, wantIdentity bool) (int64, bool) {
	vm.adodbConnectionClearErrors(conn)
	if conn.state != adStateOpen {
		vm.adodbConnectionOpen(conn)
		if conn.state != adStateOpen {
			return 0, false
		}
	}
	if conn.db != nil {
		var (
			res sql.Result
			err error
		)
		if conn.tx != nil {
			res, err = conn.tx.Exec(sqlText)
		} else {
			res, err = conn.db.Exec(sqlText)
		}
		if err != nil {
			vm.adodbConnectionRaiseProviderError(conn, source, err.Error(), "")
			return 0, false
		}
		insertID, _ := res.LastInsertId()
		return insertID, true
	}
	if conn.oleConnection != nil {
		res, err := oleutil.CallMethod(conn.oleConnection, "Execute", sqlText)
		if err != nil {
			vm.adodbConnectionRaiseProviderError(conn, source, "OLE: "+err.Error(), "")
			return 0, false
		}
		if res != nil {
			res.Clear()
		}
		if wantIdentity {
			identity, ok := vm.adodbOLEReadIdentity(conn)
			if ok {
				return identity, true
			}
		}
		return 0, true
	}
	vm.raise(vbscript.ObjectVariableNotSet, source+": no active connection")
	return 0, false
}

// adodbOLEReadIdentity reads @@IDENTITY from one live OLE connection after INSERT statements.
func (vm *VM) adodbOLEReadIdentity(conn *adodbConnection) (int64, bool) {
	if conn == nil || conn.oleConnection == nil {
		return 0, false
	}
	res, err := oleutil.CallMethod(conn.oleConnection, "Execute", "SELECT @@IDENTITY AS NewID")
	if err != nil || res == nil {
		return 0, false
	}
	disp := res.ToIDispatch()
	if disp == nil {
		res.Clear()
		return 0, false
	}
	disp.AddRef()
	res.Clear()
	defer disp.Release()
	eofRes, eofErr := oleutil.GetProperty(disp, "EOF")
	if eofErr == nil && eofRes != nil {
		eof := eofRes.Value()
		eofRes.Clear()
		if b, ok := eof.(bool); ok && b {
			return 0, false
		}
	}
	fieldsRes, fieldsErr := oleutil.GetProperty(disp, "Fields")
	if fieldsErr != nil || fieldsRes == nil {
		return 0, false
	}
	fields := fieldsRes.ToIDispatch()
	if fields != nil {
		fields.AddRef()
	}
	fieldsRes.Clear()
	if fields == nil {
		return 0, false
	}
	defer fields.Release()
	itemRes, itemErr := oleutil.CallMethod(fields, "Item", 0)
	if itemErr != nil || itemRes == nil {
		return 0, false
	}
	item := itemRes.ToIDispatch()
	if item != nil {
		item.AddRef()
	}
	itemRes.Clear()
	if item == nil {
		return 0, false
	}
	defer item.Release()
	valRes, valErr := oleutil.GetProperty(item, "Value")
	if valErr != nil || valRes == nil {
		return 0, false
	}
	defer valRes.Clear()
	identity := vm.adodbValueToVMValue(valRes.Value())
	if identity.Type == VTInteger {
		return identity.Num, true
	}
	if identity.Type == VTDouble {
		return int64(identity.Flt), true
	}
	if identity.Type == VTString {
		parsed, err := strconv.ParseInt(strings.TrimSpace(identity.Str), 10, 64)
		if err == nil {
			return parsed, true
		}
	}
	return 0, false
}

// --- Helpers ---

func (vm *VM) adodbParseConnectionString(connStr string) (driver string, dsn string) {
	trimmed := strings.TrimSpace(connStr)
	connStrLower := strings.ToLower(trimmed)

	// Direct URI prefixes bypass key=value parsing.
	if strings.HasPrefix(connStrLower, "sqlite:") {
		driver = "sqlite"
		dsn = strings.TrimPrefix(trimmed, trimmed[:len("sqlite:")])
		if dsn == "" {
			dsn = ":memory:"
		}
		return
	}
	if strings.HasPrefix(connStrLower, "oracle://") {
		driver = "oracle"
		dsn = trimmed
		return
	}

	params := adodbParseConnectionStringParams(trimmed)

	driverStr := strings.ToLower(params["driver"])
	providerStr := strings.ToLower(params["provider"])

	server := vm.adodbFirstNonEmpty(params["server"], params["data source"], params["datasource"], params["host"])
	database := vm.adodbFirstNonEmpty(params["database"], params["initial catalog"], params["dbname"], params["data source"], params["datasource"])
	uid := vm.adodbFirstNonEmpty(params["uid"], params["user id"], params["user"])
	pwd := vm.adodbFirstNonEmpty(params["pwd"], params["password"])

	if strings.Contains(driverStr, "mysql") {
		driver = "mysql"
		port := params["port"]
		if port == "" {
			port = "3306"
		}
		dsn = fmt.Sprintf("%s:%s@tcp(%s:%s)/%s?parseTime=true", uid, pwd, server, port, database)
	} else if strings.Contains(driverStr, "postgres") {
		driver = "postgres"
		port := params["port"]
		if port == "" {
			port = "5432"
		}
		dsn = fmt.Sprintf("user=%s password=%s host=%s port=%s dbname=%s sslmode=disable", uid, pwd, server, port, database)
	} else if strings.Contains(driverStr, "sql server") || strings.Contains(driverStr, "mssql") || strings.Contains(providerStr, "sqloledb") {
		driver = "mssql"
		dsn = fmt.Sprintf("server=%s;user id=%s;password=%s;database=%s", server, uid, pwd, database)
	} else if strings.Contains(driverStr, "sqlite") {
		driver = "sqlite"
		dsn = vm.adodbFirstNonEmpty(params["data source"], params["datasource"], database)
		if dsn == "" {
			dsn = ":memory:"
		}
	} else if strings.Contains(driverStr, "oracle") || strings.Contains(providerStr, "oraoledb") || strings.Contains(providerStr, "msdaora") {
		driver = "oracle"
		dsn = vm.adodbBuildOracleDSN(params, server, uid, pwd)
	}

	return
}

// adodbBuildOracleDSN constructs the oracle:// URL DSN from Classic ASP key=value connection parameters.
// It supports both explicit SID/Service Name keys and inline host:port/sid notation inside Data Source.
func (vm *VM) adodbBuildOracleDSN(params map[string]string, server, uid, pwd string) string {
	sid := vm.adodbFirstNonEmpty(params["sid"], params["service name"], params["service_name"])
	port := params["port"]

	// Parse "host:port/sid" notation that is commonly placed in the Data Source key.
	if sid == "" && strings.Contains(server, "/") {
		parts := strings.SplitN(server, "/", 2)
		hostPort := parts[0]
		sid = parts[1]
		if strings.Contains(hostPort, ":") {
			hp := strings.SplitN(hostPort, ":", 2)
			server = hp[0]
			if port == "" {
				port = hp[1]
			}
		} else {
			server = hostPort
		}
	}

	if port == "" {
		port = "1521"
	}

	// URL-encode the password so that special characters are safe inside the DSN URL.
	pwdEncoded := url.QueryEscape(pwd)
	return fmt.Sprintf("oracle://%s:%s@%s:%s/%s", uid, pwdEncoded, server, port, sid)
}

func (vm *VM) adodbFirstNonEmpty(values ...string) string {
	for _, v := range values {
		if v != "" {
			return v
		}
	}
	return ""
}

func (vm *VM) adodbIsQuery(sql string) bool {
	s := strings.ToLower(strings.TrimSpace(sql))
	return strings.HasPrefix(s, "select") || strings.HasPrefix(s, "show") || strings.HasPrefix(s, "pragma")
}

// adodbNormalizeRecordsetSource rewrites bare table names passed to Recordset.Open
// into a SELECT * FROM query while leaving explicit SQL text untouched.
func (vm *VM) adodbNormalizeRecordsetSource(sqlText string, conn *adodbConnection) string {
	trimmed := strings.TrimSpace(sqlText)
	if trimmed == "" || vm.adodbIsQuery(trimmed) || !vm.adodbLooksLikeSingleIdentifier(trimmed) {
		return sqlText
	}
	return "SELECT * FROM " + vm.adodbQuoteIdentifier(conn, trimmed)
}

// adodbLooksLikeSingleIdentifier reports whether the Recordset.Open source is a
// bare table identifier with optional schema qualification and no SQL syntax.
func (vm *VM) adodbLooksLikeSingleIdentifier(source string) bool {
	parts := strings.Split(source, ".")
	if len(parts) == 0 {
		return false
	}
	for _, part := range parts {
		if part == "" {
			return false
		}
		for _, r := range part {
			if r == '_' || r == '$' {
				continue
			}
			if unicode.IsLetter(r) || unicode.IsDigit(r) {
				continue
			}
			return false
		}
	}
	return true
}

// adodbQuoteIdentifier quotes one identifier according to the active SQL driver.
func (vm *VM) adodbQuoteIdentifier(conn *adodbConnection, identifier string) string {
	parts := strings.Split(identifier, ".")
	quoted := make([]string, 0, len(parts))
	for _, part := range parts {
		switch {
		case conn != nil && conn.dbDriver == "mysql":
			quoted = append(quoted, "`"+strings.ReplaceAll(part, "`", "``")+"`")
		case conn != nil && (conn.dbDriver == "postgres" || conn.dbDriver == "oracle"):
			quoted = append(quoted, `"`+strings.ReplaceAll(part, `"`, `""`)+`"`)
		default:
			quoted = append(quoted, "["+strings.ReplaceAll(part, "]", "]]")+"]")
		}
	}
	return strings.Join(quoted, ".")
}

func (vm *VM) adodbValueToVMValue(v interface{}) Value {
	if v == nil {
		return NewNull()
	}
	switch val := v.(type) {
	case ole.VARIANT:
		return vm.adodbValueToVMValue(val.Value())
	case *ole.VARIANT:
		if val == nil {
			return NewNull()
		}
		return vm.adodbValueToVMValue(val.Value())
	case int32:
		return NewInteger(int64(val))
	case uint32:
		return NewInteger(int64(val))
	case uint64:
		if val > math.MaxInt64 {
			return NewDouble(float64(val))
		}
		return NewInteger(int64(val))
	case int64:
		return NewInteger(val)
	case int:
		return NewInteger(int64(val))
	case int16:
		return NewInteger(int64(val))
	case int8:
		return NewInteger(int64(val))
	case uint16:
		return NewInteger(int64(val))
	case uint8:
		return NewInteger(int64(val))
	case float64:
		return NewDouble(val)
	case float32:
		return NewDouble(float64(val))
	case string:
		return NewString(val)
	case []uint16:
		return NewString(string(utf16.Decode(val)))
	case bool:
		return NewBool(val)
	case time.Time:
		return NewString(val.Format("2006-01-02 15:04:05"))
	case []byte:
		return NewString(string(val))
	}
	return NewString(fmt.Sprintf("%v", v))
}

// adodbExecuteArgs converts ADODB.Connection.Execute optional parameter payload
// to database/sql positional arguments.
// Supported forms:
//
//	Execute sql, valuesArray
//	Execute sql, singleValue
func (vm *VM) adodbExecuteArgs(args []Value) []interface{} {
	if len(args) < 2 {
		return nil
	}

	payload := args[1]
	if payload.Type == VTArray && payload.Arr != nil {
		values := payload.Arr.Values
		params := make([]interface{}, 0, len(values))
		for i := 0; i < len(values); i++ {
			params = append(params, vm.adodbDriverValue(values[i]))
		}
		return params
	}

	return []interface{}{vm.adodbDriverValue(payload)}
}

// adodbDriverValue maps VM values to driver-compatible Go values.
func (vm *VM) adodbDriverValue(v Value) interface{} {
	switch v.Type {
	case VTEmpty, VTNull:
		return nil
	case VTBool:
		return v.Num != 0
	case VTInteger:
		return v.Num
	case VTDouble:
		return v.Flt
	case VTString:
		return v.Str
	case VTDate:
		return time.Unix(0, v.Num).UTC()
	default:
		return v.String()
	}
}

// adodbApplyConnectionMetadata extracts compatibility properties from a connection string.
func (vm *VM) adodbApplyConnectionMetadata(conn *adodbConnection, connStr string) {
	if conn == nil {
		return
	}
	params := adodbParseConnectionStringParams(connStr)
	if provider := strings.TrimSpace(params["provider"]); provider != "" {
		conn.provider = provider
	}
	if conn.provider == "" {
		driver, _ := vm.adodbParseConnectionString(connStr)
		switch driver {
		case "sqlite":
			conn.provider = "SQLite"
		case "mysql":
			conn.provider = "MySQL ODBC"
		case "postgres":
			conn.provider = "PostgreSQL ODBC"
		case "mssql":
			conn.provider = "SQLOLEDB"
		}
	}
	if database := vm.adodbFirstNonEmpty(params["database"], params["initial catalog"], params["dbname"], params["data source"], params["datasource"]); database != "" {
		conn.defaultDatabase = database
	}
	if timeout := adodbParseConnectionInteger(params, "connect timeout", "connection timeout"); timeout > 0 {
		conn.connectionTimeout = timeout
	}
	if timeout := adodbParseConnectionInteger(params, "command timeout"); timeout > 0 {
		conn.commandTimeout = timeout
	}
	if cursorLocation := adodbParseConnectionInteger(params, "cursor location"); cursorLocation > 0 {
		conn.cursorLocation = cursorLocation
	}
	if isolationLevel := adodbParseConnectionInteger(params, "isolation level"); isolationLevel != 0 {
		conn.isolationLevel = isolationLevel
	}
}

// adodbSchemaTablesQuery returns the provider-specific SQL used by OpenSchema.
func (vm *VM) adodbSchemaTablesQuery(driver string) string {
	switch driver {
	case "sqlite":
		return "SELECT name, upper(type) FROM sqlite_master WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' ORDER BY name"
	case "mysql", "postgres", "mssql":
		return "SELECT table_name, table_type FROM information_schema.tables ORDER BY table_name"
	case "oracle":
		return "SELECT table_name, 'TABLE' AS table_type FROM user_tables UNION ALL SELECT view_name, 'VIEW' AS table_type FROM user_views ORDER BY 1"
	default:
		return ""
	}
}

// adodbSchemaColumnNames returns the projected column list for one ADO schema enum.
func (vm *VM) adodbSchemaColumnNames(schemaID int) []string {
	switch schemaID {
	case adSchemaProcedures:
		return []string{"PROCEDURE_NAME", "PROCEDURE_TYPE"}
	case adSchemaViews:
		return []string{"TABLE_NAME"}
	case adSchemaColumns:
		return []string{"TABLE_NAME", "COLUMN_NAME", "ORDINAL_POSITION", "DATA_TYPE", "CHARACTER_MAXIMUM_LENGTH", "NUMERIC_SCALE", "IS_NULLABLE"}
	case adSchemaIndexes:
		return []string{"TABLE_NAME", "INDEX_NAME", "PRIMARY_KEY", "UNIQUE", "TYPE", "ORDINAL_POSITION", "COLUMN_NAME"}
	case adSchemaForeignKeys:
		return []string{"PK_TABLE_NAME", "PK_COLUMN_NAME", "FK_TABLE_NAME", "FK_COLUMN_NAME", "KEY_SEQ"}
	default:
		return []string{"TABLE_NAME", "TABLE_TYPE"}
	}
}

// adodbSchemaRestrictions reads one OpenSchema restrictions array.
func (vm *VM) adodbSchemaRestrictions(args []Value) []string {
	if len(args) < 2 || args[1].Type != VTArray || args[1].Arr == nil {
		return nil
	}
	restrictions := make([]string, len(args[1].Arr.Values))
	for i := 0; i < len(args[1].Arr.Values); i++ {
		restrictions[i] = strings.TrimSpace(vm.valueToString(args[1].Arr.Values[i]))
	}
	return restrictions
}

// adodbSchemaRestriction returns one restriction slot or an empty string when omitted.
func (vm *VM) adodbSchemaRestriction(restrictions []string, index int) string {
	if index < 0 || index >= len(restrictions) {
		return ""
	}
	return strings.TrimSpace(restrictions[index])
}

// adodbBuildSchemaRows materializes one OpenSchema result set for native DB providers.
func (vm *VM) adodbBuildSchemaRows(conn *adodbConnection, schemaID int, restrictions []string) []map[string]Value {
	if conn == nil || conn.db == nil {
		return nil
	}
	switch schemaID {
	case adSchemaProcedures:
		return vm.adodbBuildProceduresSchemaRows(conn, restrictions)
	case adSchemaViews:
		return vm.adodbBuildViewsSchemaRows(conn, restrictions)
	case adSchemaColumns:
		return vm.adodbBuildColumnsSchemaRows(conn, restrictions)
	case adSchemaIndexes:
		return vm.adodbBuildIndexesSchemaRows(conn, restrictions)
	case adSchemaForeignKeys:
		return vm.adodbBuildForeignKeysSchemaRows(conn, restrictions)
	default:
		return vm.adodbBuildTablesSchemaRows(conn, restrictions)
	}
}

// adodbBuildTablesSchemaRows returns table/view rows honoring ADO table restrictions.
func (vm *VM) adodbBuildTablesSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := vm.adodbSchemaTablesQuery(conn.dbDriver)
	if query == "" {
		return nil
	}
	rows, err := conn.db.Query(query)
	if err != nil {
		return nil
	}
	defer rows.Close()

	tableFilter := vm.adodbSchemaRestriction(restrictions, 2)
	typeFilter := strings.ToUpper(vm.adodbSchemaRestriction(restrictions, 3))
	results := make([]map[string]Value, 0, 8)
	for rows.Next() {
		var tableName string
		var tableType string
		if scanErr := rows.Scan(&tableName, &tableType); scanErr != nil {
			continue
		}
		if tableFilter != "" && !strings.EqualFold(tableName, tableFilter) {
			continue
		}
		normalizedType := vm.adodbNormalizeSchemaTableType(tableType)
		if typeFilter != "" && normalizedType != typeFilter {
			continue
		}
		results = append(results, map[string]Value{
			"table_name": NewString(tableName),
			"table_type": NewString(normalizedType),
		})
	}
	return results
}

// adodbBuildProceduresSchemaRows returns procedure metadata honoring ADO procedure restrictions.
func (vm *VM) adodbBuildProceduresSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	switch conn.dbDriver {
	case "mysql", "postgres", "mssql":
		return vm.adodbBuildInformationSchemaProceduresRows(conn, restrictions)
	case "oracle":
		return vm.adodbBuildOracleProceduresSchemaRows(conn, restrictions)
	default:
		return nil
	}
}

// adodbBuildViewsSchemaRows returns view metadata honoring ADO view restrictions.
func (vm *VM) adodbBuildViewsSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	viewFilter := vm.adodbSchemaRestriction(restrictions, 2)
	switch conn.dbDriver {
	case "sqlite":
		tables := vm.adodbBuildTablesSchemaRows(conn, []string{"", "", viewFilter, "VIEW"})
		results := make([]map[string]Value, 0, len(tables))
		for i := 0; i < len(tables); i++ {
			results = append(results, map[string]Value{
				"table_name": NewString(tables[i]["table_name"].String()),
			})
		}
		return results
	case "oracle":
		return vm.adodbBuildOracleViewsSchemaRows(conn, restrictions)
	default:
		return vm.adodbBuildInformationSchemaViewsRows(conn, restrictions)
	}
}

// adodbBuildColumnsSchemaRows returns column metadata honoring ADO column restrictions.
func (vm *VM) adodbBuildColumnsSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	switch conn.dbDriver {
	case "sqlite":
		return vm.adodbBuildSQLiteColumnsSchemaRows(conn, restrictions)
	case "oracle":
		return vm.adodbBuildOracleColumnsSchemaRows(conn, restrictions)
	default:
		return vm.adodbBuildInformationSchemaColumnsRows(conn, restrictions)
	}
}

// adodbBuildForeignKeysSchemaRows returns foreign key metadata honoring ADO foreign-key restrictions.
func (vm *VM) adodbBuildForeignKeysSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	switch conn.dbDriver {
	case "sqlite":
		return vm.adodbBuildSQLiteForeignKeysSchemaRows(conn, restrictions)
	case "mysql", "postgres", "mssql":
		return vm.adodbBuildInformationSchemaForeignKeysRows(conn, restrictions)
	case "oracle":
		return vm.adodbBuildOracleForeignKeysSchemaRows(conn, restrictions)
	default:
		return nil
	}
}

// adodbBuildInformationSchemaProceduresRows uses INFORMATION_SCHEMA.ROUTINES for procedure metadata.
func (vm *VM) adodbBuildInformationSchemaProceduresRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := "SELECT routine_name, routine_type FROM information_schema.routines ORDER BY routine_name"
	rows, err := conn.db.Query(query)
	if err != nil {
		return nil
	}
	defer rows.Close()

	procedureFilter := vm.adodbSchemaRestriction(restrictions, 2)
	typeFilter := strings.ToUpper(vm.adodbSchemaRestriction(restrictions, 3))
	results := make([]map[string]Value, 0, 8)
	for rows.Next() {
		var procedureName string
		var procedureType string
		if scanErr := rows.Scan(&procedureName, &procedureType); scanErr != nil {
			continue
		}
		if procedureFilter != "" && !strings.EqualFold(procedureName, procedureFilter) {
			continue
		}
		normalizedType := strings.ToUpper(strings.TrimSpace(procedureType))
		if typeFilter != "" && normalizedType != typeFilter {
			continue
		}
		results = append(results, map[string]Value{
			"procedure_name": NewString(procedureName),
			"procedure_type": NewString(normalizedType),
		})
	}
	return results
}

// adodbBuildInformationSchemaViewsRows uses INFORMATION_SCHEMA.VIEWS for view metadata.
func (vm *VM) adodbBuildInformationSchemaViewsRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := "SELECT table_name FROM information_schema.views ORDER BY table_name"
	rows, err := conn.db.Query(query)
	if err != nil {
		return nil
	}
	defer rows.Close()

	viewFilter := vm.adodbSchemaRestriction(restrictions, 2)
	results := make([]map[string]Value, 0, 8)
	for rows.Next() {
		var viewName string
		if scanErr := rows.Scan(&viewName); scanErr != nil {
			continue
		}
		if viewFilter != "" && !strings.EqualFold(viewName, viewFilter) {
			continue
		}
		results = append(results, map[string]Value{
			"table_name": NewString(viewName),
		})
	}
	return results
}

// adodbBuildIndexesSchemaRows returns index metadata honoring ADO index restrictions.
func (vm *VM) adodbBuildIndexesSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	switch conn.dbDriver {
	case "sqlite":
		return vm.adodbBuildSQLiteIndexesSchemaRows(conn, restrictions)
	case "mysql":
		return vm.adodbBuildMySQLIndexesSchemaRows(conn, restrictions)
	case "postgres":
		return vm.adodbBuildPostgresIndexesSchemaRows(conn, restrictions)
	case "mssql":
		return vm.adodbBuildMSSQLIndexesSchemaRows(conn, restrictions)
	case "oracle":
		return vm.adodbBuildOracleIndexesSchemaRows(conn, restrictions)
	default:
		return nil
	}
}

// adodbBuildSQLiteColumnsSchemaRows uses PRAGMA table_info for SQLite column metadata.
func (vm *VM) adodbBuildSQLiteColumnsSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	tableFilter := vm.adodbSchemaRestriction(restrictions, 2)
	columnFilter := vm.adodbSchemaRestriction(restrictions, 3)
	tables := vm.adodbBuildTablesSchemaRows(conn, []string{"", "", tableFilter, ""})
	results := make([]map[string]Value, 0, 16)
	for i := 0; i < len(tables); i++ {
		tableName := tables[i]["table_name"].String()
		pragma := "PRAGMA table_info('" + strings.ReplaceAll(tableName, "'", "''") + "')"
		rows, err := conn.db.Query(pragma)
		if err != nil {
			continue
		}
		for rows.Next() {
			var columnID int
			var columnName string
			var dataType string
			var notNull int
			var defaultValue sql.NullString
			var primaryKey int
			if scanErr := rows.Scan(&columnID, &columnName, &dataType, &notNull, &defaultValue, &primaryKey); scanErr != nil {
				continue
			}
			if columnFilter != "" && !strings.EqualFold(columnName, columnFilter) {
				continue
			}
			charLen, scale := adodbSQLiteTypeMetadata(dataType)
			isNullable := "YES"
			if notNull != 0 || primaryKey != 0 {
				isNullable = "NO"
			}
			results = append(results, map[string]Value{
				"table_name":               NewString(tableName),
				"column_name":              NewString(columnName),
				"ordinal_position":         NewInteger(int64(columnID + 1)),
				"data_type":                NewString(strings.TrimSpace(dataType)),
				"character_maximum_length": NewInteger(int64(charLen)),
				"numeric_scale":            NewInteger(int64(scale)),
				"is_nullable":              NewString(isNullable),
			})
		}
		_ = rows.Close()
	}
	return results
}

// adodbBuildInformationSchemaColumnsRows uses INFORMATION_SCHEMA.COLUMNS for SQL providers.
func (vm *VM) adodbBuildInformationSchemaColumnsRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := "SELECT table_name, column_name, ordinal_position, data_type, COALESCE(character_maximum_length, numeric_precision, 0), COALESCE(numeric_scale, 0), COALESCE(is_nullable, 'YES') FROM information_schema.columns ORDER BY table_name, ordinal_position"
	rows, err := conn.db.Query(query)
	if err != nil {
		return nil
	}
	defer rows.Close()

	tableFilter := vm.adodbSchemaRestriction(restrictions, 2)
	columnFilter := vm.adodbSchemaRestriction(restrictions, 3)
	results := make([]map[string]Value, 0, 16)
	for rows.Next() {
		var tableName string
		var columnName string
		var ordinalPosition int64
		var dataType string
		var charLen sql.NullInt64
		var numericScale sql.NullInt64
		var isNullable sql.NullString
		if scanErr := rows.Scan(&tableName, &columnName, &ordinalPosition, &dataType, &charLen, &numericScale, &isNullable); scanErr != nil {
			continue
		}
		if tableFilter != "" && !strings.EqualFold(tableName, tableFilter) {
			continue
		}
		if columnFilter != "" && !strings.EqualFold(columnName, columnFilter) {
			continue
		}
		results = append(results, map[string]Value{
			"table_name":               NewString(tableName),
			"column_name":              NewString(columnName),
			"ordinal_position":         NewInteger(ordinalPosition),
			"data_type":                NewString(dataType),
			"character_maximum_length": NewInteger(adodbNullInt64(charLen)),
			"numeric_scale":            NewInteger(adodbNullInt64(numericScale)),
			"is_nullable":              NewString(adodbNullStringOrDefault(isNullable, "YES")),
		})
	}
	return results
}

// adodbBuildSQLiteIndexesSchemaRows uses PRAGMA index_list/index_info for SQLite index metadata.
// The nested PRAGMA index_info queries are done AFTER closing the PRAGMA index_list cursor to avoid
// a SQLite connection-pool deadlock: database/sql may reuse the same underlying connection and
// SQLite cannot run a second statement while a cursor on the first is still open.
func (vm *VM) adodbBuildSQLiteIndexesSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	indexFilter := vm.adodbSchemaRestriction(restrictions, 2)
	typeFilter := strings.TrimSpace(vm.adodbSchemaRestriction(restrictions, 3))
	tableFilter := vm.adodbSchemaRestriction(restrictions, 4)
	tables := vm.adodbBuildTablesSchemaRows(conn, []string{"", "", tableFilter, "TABLE"})
	results := make([]map[string]Value, 0, 16)

	// indexEntry holds one row from PRAGMA index_list before we close that cursor.
	type indexEntry struct {
		tableName string
		indexName string
		unique    int
		origin    string
	}

	for i := 0; i < len(tables); i++ {
		tableName := tables[i]["table_name"].String()
		listQuery := "PRAGMA index_list('" + strings.ReplaceAll(tableName, "'", "''") + "')"
		listRows, err := conn.db.Query(listQuery)
		if err != nil {
			continue
		}

		// Step 1: collect all index list entries into memory, then close the cursor.
		// This prevents a nested-query deadlock when index_info is queried below.
		var entries []indexEntry
		for listRows.Next() {
			var seq int
			var indexName string
			var unique int
			var origin string
			var partial int
			if scanErr := listRows.Scan(&seq, &indexName, &unique, &origin, &partial); scanErr != nil {
				continue
			}
			if indexFilter != "" && !strings.EqualFold(indexName, indexFilter) {
				continue
			}
			indexType := vm.adodbSQLiteIndexType(origin)
			if typeFilter != "" && !strings.EqualFold(indexType, typeFilter) {
				continue
			}
			entries = append(entries, indexEntry{tableName: tableName, indexName: indexName, unique: unique, origin: origin})
		}
		_ = listRows.Close()

		// Step 2: query index_info for each collected entry — cursor is fully closed above.
		for _, entry := range entries {
			infoQuery := "PRAGMA index_info('" + strings.ReplaceAll(entry.indexName, "'", "''") + "')"
			infoRows, infoErr := conn.db.Query(infoQuery)
			if infoErr != nil {
				continue
			}
			indexType := vm.adodbSQLiteIndexType(entry.origin)
			for infoRows.Next() {
				var columnSeq int
				var ordinalPosition int
				var columnName string
				if scanErr := infoRows.Scan(&columnSeq, &ordinalPosition, &columnName); scanErr != nil {
					continue
				}
				results = append(results, map[string]Value{
					"table_name":       NewString(entry.tableName),
					"index_name":       NewString(entry.indexName),
					"primary_key":      NewBool(strings.EqualFold(entry.origin, "pk")),
					"unique":           NewBool(entry.unique != 0),
					"type":             NewString(indexType),
					"ordinal_position": NewInteger(int64(ordinalPosition + 1)),
					"column_name":      NewString(columnName),
				})
			}
			_ = infoRows.Close()
		}
	}
	return results
}

// adodbBuildSQLiteForeignKeysSchemaRows uses PRAGMA foreign_key_list for SQLite foreign-key metadata.
func (vm *VM) adodbBuildSQLiteForeignKeysSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	pkTableFilter := vm.adodbSchemaRestriction(restrictions, 2)
	fkTableFilter := vm.adodbSchemaRestriction(restrictions, 5)
	tables := vm.adodbBuildTablesSchemaRows(conn, []string{"", "", fkTableFilter, "TABLE"})
	results := make([]map[string]Value, 0, 8)
	for i := 0; i < len(tables); i++ {
		fkTableName := tables[i]["table_name"].String()
		pragma := "PRAGMA foreign_key_list('" + strings.ReplaceAll(fkTableName, "'", "''") + "')"
		rows, err := conn.db.Query(pragma)
		if err != nil {
			continue
		}
		for rows.Next() {
			var id int
			var seq int
			var pkTableName string
			var fkColumnName string
			var pkColumnName string
			var onUpdate string
			var onDelete string
			var match string
			if scanErr := rows.Scan(&id, &seq, &pkTableName, &fkColumnName, &pkColumnName, &onUpdate, &onDelete, &match); scanErr != nil {
				continue
			}
			if pkTableFilter != "" && !strings.EqualFold(pkTableName, pkTableFilter) {
				continue
			}
			results = append(results, map[string]Value{
				"pk_table_name":  NewString(pkTableName),
				"pk_column_name": NewString(pkColumnName),
				"fk_table_name":  NewString(fkTableName),
				"fk_column_name": NewString(fkColumnName),
				"key_seq":        NewInteger(int64(seq + 1)),
			})
		}
		_ = rows.Close()
	}
	return results
}

// adodbBuildMySQLIndexesSchemaRows uses information_schema.statistics for MySQL index metadata.
func (vm *VM) adodbBuildMySQLIndexesSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := "SELECT table_name, index_name, non_unique, index_type, seq_in_index, column_name FROM information_schema.statistics ORDER BY table_name, index_name, seq_in_index"
	return vm.adodbBuildInformationSchemaIndexRows(conn, query, restrictions)
}

// adodbBuildPostgresIndexesSchemaRows uses pg_catalog joins for PostgreSQL index metadata.
func (vm *VM) adodbBuildPostgresIndexesSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := "SELECT t.relname AS table_name, i.relname AS index_name, CASE WHEN ix.indisprimary THEN 1 ELSE 0 END AS primary_key, CASE WHEN ix.indisunique THEN 1 ELSE 0 END AS is_unique, am.amname AS index_type, (ord.ordinality) AS seq_in_index, a.attname AS column_name FROM pg_class t JOIN pg_index ix ON t.oid = ix.indrelid JOIN pg_class i ON i.oid = ix.indexrelid JOIN pg_am am ON i.relam = am.oid JOIN LATERAL unnest(ix.indkey) WITH ORDINALITY AS ord(attnum, ordinality) ON TRUE JOIN pg_attribute a ON a.attrelid = t.oid AND a.attnum = ord.attnum WHERE t.relkind = 'r' ORDER BY t.relname, i.relname, ord.ordinality"
	return vm.adodbBuildCatalogIndexRows(conn, query, restrictions, true)
}

// adodbBuildMSSQLIndexesSchemaRows uses sys catalog views for SQL Server index metadata.
func (vm *VM) adodbBuildMSSQLIndexesSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := "SELECT t.name AS table_name, i.name AS index_name, CASE WHEN i.is_primary_key = 1 THEN 1 ELSE 0 END AS primary_key, CASE WHEN i.is_unique = 1 THEN 1 ELSE 0 END AS is_unique, i.type_desc AS index_type, ic.key_ordinal AS seq_in_index, c.name AS column_name FROM sys.indexes i JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id JOIN sys.tables t ON i.object_id = t.object_id WHERE i.name IS NOT NULL ORDER BY t.name, i.name, ic.key_ordinal"
	return vm.adodbBuildCatalogIndexRows(conn, query, restrictions, false)
}

// adodbBuildInformationSchemaIndexRows maps INFORMATION_SCHEMA index metadata into OpenSchema rows.
func (vm *VM) adodbBuildInformationSchemaIndexRows(conn *adodbConnection, query string, restrictions []string) []map[string]Value {
	rows, err := conn.db.Query(query)
	if err != nil {
		return nil
	}
	defer rows.Close()
	indexFilter := vm.adodbSchemaRestriction(restrictions, 2)
	typeFilter := vm.adodbSchemaRestriction(restrictions, 3)
	tableFilter := vm.adodbSchemaRestriction(restrictions, 4)
	results := make([]map[string]Value, 0, 16)
	for rows.Next() {
		var tableName string
		var indexName string
		var nonUnique int64
		var indexType string
		var seqInIndex int64
		var columnName string
		if scanErr := rows.Scan(&tableName, &indexName, &nonUnique, &indexType, &seqInIndex, &columnName); scanErr != nil {
			continue
		}
		if tableFilter != "" && !strings.EqualFold(tableName, tableFilter) {
			continue
		}
		if indexFilter != "" && !strings.EqualFold(indexName, indexFilter) {
			continue
		}
		if typeFilter != "" && !strings.EqualFold(indexType, typeFilter) {
			continue
		}
		results = append(results, map[string]Value{
			"table_name":       NewString(tableName),
			"index_name":       NewString(indexName),
			"primary_key":      NewBool(strings.EqualFold(indexName, "PRIMARY")),
			"unique":           NewBool(nonUnique == 0),
			"type":             NewString(indexType),
			"ordinal_position": NewInteger(seqInIndex),
			"column_name":      NewString(columnName),
		})
	}
	return results
}

// adodbBuildCatalogIndexRows maps provider catalog query output into OpenSchema rows.
func (vm *VM) adodbBuildCatalogIndexRows(conn *adodbConnection, query string, restrictions []string, lowercaseType bool) []map[string]Value {
	rows, err := conn.db.Query(query)
	if err != nil {
		return nil
	}
	defer rows.Close()
	indexFilter := vm.adodbSchemaRestriction(restrictions, 2)
	typeFilter := vm.adodbSchemaRestriction(restrictions, 3)
	tableFilter := vm.adodbSchemaRestriction(restrictions, 4)
	results := make([]map[string]Value, 0, 16)
	for rows.Next() {
		var tableName string
		var indexName string
		var primaryKey int64
		var isUnique int64
		var indexType string
		var seqInIndex int64
		var columnName string
		if scanErr := rows.Scan(&tableName, &indexName, &primaryKey, &isUnique, &indexType, &seqInIndex, &columnName); scanErr != nil {
			continue
		}
		if tableFilter != "" && !strings.EqualFold(tableName, tableFilter) {
			continue
		}
		if indexFilter != "" && !strings.EqualFold(indexName, indexFilter) {
			continue
		}
		normalizedType := strings.TrimSpace(indexType)
		if lowercaseType {
			normalizedType = strings.ToLower(normalizedType)
		}
		if typeFilter != "" && !strings.EqualFold(normalizedType, typeFilter) {
			continue
		}
		results = append(results, map[string]Value{
			"table_name":       NewString(tableName),
			"index_name":       NewString(indexName),
			"primary_key":      NewBool(primaryKey != 0),
			"unique":           NewBool(isUnique != 0),
			"type":             NewString(normalizedType),
			"ordinal_position": NewInteger(seqInIndex),
			"column_name":      NewString(columnName),
		})
	}
	return results
}

// adodbBuildInformationSchemaForeignKeysRows uses INFORMATION_SCHEMA constraint views for foreign-key metadata.
func (vm *VM) adodbBuildInformationSchemaForeignKeysRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := "SELECT pk.table_name AS pk_table_name, pk.column_name AS pk_column_name, fk.table_name AS fk_table_name, fk.column_name AS fk_column_name, fk.ordinal_position AS key_seq FROM information_schema.referential_constraints rc JOIN information_schema.key_column_usage fk ON rc.constraint_name = fk.constraint_name AND rc.constraint_schema = fk.constraint_schema JOIN information_schema.key_column_usage pk ON rc.unique_constraint_name = pk.constraint_name AND rc.unique_constraint_schema = pk.constraint_schema AND fk.ordinal_position = pk.ordinal_position ORDER BY fk.table_name, fk.ordinal_position"
	rows, err := conn.db.Query(query)
	if err != nil {
		return nil
	}
	defer rows.Close()

	pkTableFilter := vm.adodbSchemaRestriction(restrictions, 2)
	fkTableFilter := vm.adodbSchemaRestriction(restrictions, 5)
	results := make([]map[string]Value, 0, 16)
	for rows.Next() {
		var pkTableName string
		var pkColumnName string
		var fkTableName string
		var fkColumnName string
		var keySeq int64
		if scanErr := rows.Scan(&pkTableName, &pkColumnName, &fkTableName, &fkColumnName, &keySeq); scanErr != nil {
			continue
		}
		if pkTableFilter != "" && !strings.EqualFold(pkTableName, pkTableFilter) {
			continue
		}
		if fkTableFilter != "" && !strings.EqualFold(fkTableName, fkTableFilter) {
			continue
		}
		results = append(results, map[string]Value{
			"pk_table_name":  NewString(pkTableName),
			"pk_column_name": NewString(pkColumnName),
			"fk_table_name":  NewString(fkTableName),
			"fk_column_name": NewString(fkColumnName),
			"key_seq":        NewInteger(keySeq),
		})
	}
	return results
}

// adodbBuildOracleProceduresSchemaRows queries USER_PROCEDURES for Oracle procedure metadata.
func (vm *VM) adodbBuildOracleProceduresSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := "SELECT object_name, object_type FROM user_procedures WHERE object_type IN ('PROCEDURE','FUNCTION') ORDER BY object_name"
	rows, err := conn.db.Query(query)
	if err != nil {
		return nil
	}
	defer rows.Close()

	procedureFilter := vm.adodbSchemaRestriction(restrictions, 2)
	typeFilter := strings.ToUpper(vm.adodbSchemaRestriction(restrictions, 3))
	results := make([]map[string]Value, 0, 8)
	for rows.Next() {
		var procedureName string
		var procedureType string
		if scanErr := rows.Scan(&procedureName, &procedureType); scanErr != nil {
			continue
		}
		if procedureFilter != "" && !strings.EqualFold(procedureName, procedureFilter) {
			continue
		}
		normalizedType := strings.ToUpper(strings.TrimSpace(procedureType))
		if typeFilter != "" && normalizedType != typeFilter {
			continue
		}
		results = append(results, map[string]Value{
			"procedure_name": NewString(procedureName),
			"procedure_type": NewString(normalizedType),
		})
	}
	return results
}

// adodbBuildOracleViewsSchemaRows queries USER_VIEWS for Oracle view metadata.
func (vm *VM) adodbBuildOracleViewsSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := "SELECT view_name FROM user_views ORDER BY view_name"
	rows, err := conn.db.Query(query)
	if err != nil {
		return nil
	}
	defer rows.Close()

	viewFilter := vm.adodbSchemaRestriction(restrictions, 2)
	results := make([]map[string]Value, 0, 8)
	for rows.Next() {
		var viewName string
		if scanErr := rows.Scan(&viewName); scanErr != nil {
			continue
		}
		if viewFilter != "" && !strings.EqualFold(viewName, viewFilter) {
			continue
		}
		results = append(results, map[string]Value{
			"table_name": NewString(viewName),
		})
	}
	return results
}

// adodbBuildOracleColumnsSchemaRows queries USER_TAB_COLUMNS for Oracle column metadata.
func (vm *VM) adodbBuildOracleColumnsSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := "SELECT table_name, column_name, column_id, data_type, NVL(char_length, NVL(data_precision, 0)), NVL(data_scale, 0), nullable FROM user_tab_columns ORDER BY table_name, column_id"
	rows, err := conn.db.Query(query)
	if err != nil {
		return nil
	}
	defer rows.Close()

	tableFilter := vm.adodbSchemaRestriction(restrictions, 2)
	columnFilter := vm.adodbSchemaRestriction(restrictions, 3)
	results := make([]map[string]Value, 0, 16)
	for rows.Next() {
		var tableName string
		var columnName string
		var ordinalPosition int64
		var dataType string
		var charLen int64
		var numericScale int64
		var nullable string
		if scanErr := rows.Scan(&tableName, &columnName, &ordinalPosition, &dataType, &charLen, &numericScale, &nullable); scanErr != nil {
			continue
		}
		if tableFilter != "" && !strings.EqualFold(tableName, tableFilter) {
			continue
		}
		if columnFilter != "" && !strings.EqualFold(columnName, columnFilter) {
			continue
		}
		isNullable := "YES"
		if strings.EqualFold(strings.TrimSpace(nullable), "N") {
			isNullable = "NO"
		}
		results = append(results, map[string]Value{
			"table_name":               NewString(tableName),
			"column_name":              NewString(columnName),
			"ordinal_position":         NewInteger(ordinalPosition),
			"data_type":                NewString(dataType),
			"character_maximum_length": NewInteger(charLen),
			"numeric_scale":            NewInteger(numericScale),
			"is_nullable":              NewString(isNullable),
		})
	}
	return results
}

// adodbBuildOracleIndexesSchemaRows queries USER_INDEXES and USER_IND_COLUMNS for Oracle index metadata.
func (vm *VM) adodbBuildOracleIndexesSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := `SELECT i.table_name, i.index_name,
		CASE WHEN c.constraint_type = 'P' THEN 1 ELSE 0 END AS primary_key,
		CASE WHEN i.uniqueness = 'UNIQUE' THEN 1 ELSE 0 END AS is_unique,
		i.index_type, ic.column_position, ic.column_name
		FROM user_indexes i
		JOIN user_ind_columns ic ON i.index_name = ic.index_name AND i.table_name = ic.table_name
		LEFT JOIN user_constraints c ON c.index_name = i.index_name AND c.table_name = i.table_name AND c.constraint_type = 'P'
		ORDER BY i.table_name, i.index_name, ic.column_position`
	return vm.adodbBuildCatalogIndexRows(conn, query, restrictions, false)
}

// adodbBuildOracleForeignKeysSchemaRows queries USER_CONSTRAINTS and USER_CONS_COLUMNS for Oracle foreign-key metadata.
func (vm *VM) adodbBuildOracleForeignKeysSchemaRows(conn *adodbConnection, restrictions []string) []map[string]Value {
	query := `SELECT p.table_name AS pk_table_name, pc.column_name AS pk_column_name,
		f.table_name AS fk_table_name, fc.column_name AS fk_column_name, fc.position AS key_seq
		FROM user_constraints f
		JOIN user_cons_columns fc ON f.constraint_name = fc.constraint_name
		JOIN user_constraints p ON f.r_constraint_name = p.constraint_name
		JOIN user_cons_columns pc ON p.constraint_name = pc.constraint_name AND fc.position = pc.position
		WHERE f.constraint_type = 'R'
		ORDER BY f.table_name, fc.position`
	rows, err := conn.db.Query(query)
	if err != nil {
		return nil
	}
	defer rows.Close()

	pkTableFilter := vm.adodbSchemaRestriction(restrictions, 2)
	fkTableFilter := vm.adodbSchemaRestriction(restrictions, 5)
	results := make([]map[string]Value, 0, 16)
	for rows.Next() {
		var pkTableName string
		var pkColumnName string
		var fkTableName string
		var fkColumnName string
		var keySeq int64
		if scanErr := rows.Scan(&pkTableName, &pkColumnName, &fkTableName, &fkColumnName, &keySeq); scanErr != nil {
			continue
		}
		if pkTableFilter != "" && !strings.EqualFold(pkTableName, pkTableFilter) {
			continue
		}
		if fkTableFilter != "" && !strings.EqualFold(fkTableName, fkTableFilter) {
			continue
		}
		results = append(results, map[string]Value{
			"pk_table_name":  NewString(pkTableName),
			"pk_column_name": NewString(pkColumnName),
			"fk_table_name":  NewString(fkTableName),
			"fk_column_name": NewString(fkColumnName),
			"key_seq":        NewInteger(keySeq),
		})
	}
	return results
}

// adodbHasNonEmptyRestriction reports whether one OpenSchema restrictions array carries any effective filter.
func adodbHasNonEmptyRestriction(restrictions []string) bool {
	for i := 0; i < len(restrictions); i++ {
		if strings.TrimSpace(restrictions[i]) != "" {
			return true
		}
	}
	return false
}

// adodbNormalizeSchemaTableType normalizes provider-specific table types into the ADO-style TABLE/VIEW values.
func (vm *VM) adodbNormalizeSchemaTableType(tableType string) string {
	normalizedType := strings.ToUpper(strings.TrimSpace(tableType))
	switch normalizedType {
	case "BASE TABLE", "TABLE":
		return "TABLE"
	case "VIEW", "SYSTEM VIEW":
		return "VIEW"
	default:
		return normalizedType
	}
}

// adodbSQLiteTypeMetadata extracts basic length and scale hints from a SQLite type declaration.
func adodbSQLiteTypeMetadata(typeDecl string) (int, int) {
	trimmed := strings.TrimSpace(typeDecl)
	if trimmed == "" {
		return 0, 0
	}
	openIdx := strings.Index(trimmed, "(")
	closeIdx := strings.LastIndex(trimmed, ")")
	if openIdx < 0 || closeIdx <= openIdx {
		return 0, 0
	}
	parts := strings.Split(trimmed[openIdx+1:closeIdx], ",")
	charLen := 0
	scale := 0
	if len(parts) >= 1 {
		if parsed, err := strconv.Atoi(strings.TrimSpace(parts[0])); err == nil {
			charLen = parsed
		}
	}
	if len(parts) >= 2 {
		if parsed, err := strconv.Atoi(strings.TrimSpace(parts[1])); err == nil {
			scale = parsed
		}
	}
	return charLen, scale
}

// adodbSQLiteIndexType normalizes SQLite PRAGMA index origin values to one schema type string.
func (vm *VM) adodbSQLiteIndexType(origin string) string {
	switch strings.ToLower(strings.TrimSpace(origin)) {
	case "pk":
		return "PRIMARY"
	case "u":
		return "UNIQUE"
	default:
		return "INDEX"
	}
}

// adodbNullInt64 converts nullable SQL integers to VM-compatible integer values.
func adodbNullInt64(value sql.NullInt64) int64 {
	if !value.Valid {
		return 0
	}
	return value.Int64
}

// adodbNullStringOrDefault converts nullable SQL strings with one default fallback.
func adodbNullStringOrDefault(value sql.NullString, defaultValue string) string {
	if !value.Valid {
		return defaultValue
	}
	return value.String
}

// adodbFinalizeDisconnectedRecordset normalizes disconnected recordset cursors after data changes.
func (vm *VM) adodbFinalizeDisconnectedRecordset(rs *adodbRecordset) {
	if rs == nil {
		return
	}
	rs.recordCount = len(rs.data)
	rs.state = adStateOpen
	rs.status = adRecOK
	rs.editMode = adEditNone
	if rs.recordCount == 0 {
		rs.currentRow = -1
		rs.bookmark = 0
		rs.eof = true
		rs.bof = true
		return
	}
	rs.currentRow = 0
	rs.bookmark = 1
	rs.eof = false
	rs.bof = false
}

// adodbRecordsetCompareBookmarks implements the classic ADO bookmark comparison contract.
func (vm *VM) adodbRecordsetCompareBookmarks(args []Value) Value {
	if len(args) < 2 {
		return NewInteger(0)
	}
	left := vm.asInt(args[0])
	right := vm.asInt(args[1])
	switch {
	case left < right:
		return NewInteger(-1)
	case left > right:
		return NewInteger(1)
	default:
		return NewInteger(0)
	}
}

// adodbRecordsetRequery reopens the recordset using the last recorded source.
func (vm *VM) adodbRecordsetRequery(rs *adodbRecordset) {
	if rs == nil {
		return
	}
	if rs.activeCommand != 0 {
		if cmd, ok := vm.adodbCommandItems[rs.activeCommand]; ok {
			rs.source = cmd.commandText
			if cmd.activeConnection != 0 {
				rs.activeConnection = cmd.activeConnection
			}
		}
	}
	if strings.TrimSpace(rs.source) == "" {
		vm.adodbFinalizeDisconnectedRecordset(rs)
		return
	}
	var conn *adodbConnection
	if rs.activeConnection != 0 {
		conn = vm.adodbConnectionItems[rs.activeConnection]
	}
	if conn == nil {
		vm.adodbFinalizeDisconnectedRecordset(rs)
		return
	}
	vm.adodbRecordsetOpen(rs, rs.source, conn, nil)
}

// adodbRecordsetResync keeps disconnected semantics by requerying when possible.
func (vm *VM) adodbRecordsetResync(rs *adodbRecordset) {
	vm.adodbRecordsetRequery(rs)
}

// adodbRecordsetSave persists one disconnected snapshot to a file path or ADODB.Stream.
func (vm *VM) adodbRecordsetSave(rs *adodbRecordset, args []Value) {
	if rs == nil || len(args) == 0 {
		return
	}
	payload := []byte(vm.adodbRecordsetPersistXML(rs))
	if args[0].Type == VTNativeObject {
		if stream, ok := vm.adodbStreamItems[args[0].Num]; ok {
			stream.state = adodbStateOpen
			stream.typ = adodbTypeText
			stream.position = 0
			stream.buffer = append(stream.buffer[:0], payload...)
			stream.size = int64(len(payload))
			return
		}
	}
	resolvedPath, ok := vm.adodbResolvePersistPath(args[0].String())
	if !ok {
		return
	}
	_ = os.MkdirAll(filepath.Dir(resolvedPath), 0755)
	_ = os.WriteFile(resolvedPath, payload, 0644)
}

// adodbRecordsetSeek supports bookmark seeks and simple indexed key lookup.
func (vm *VM) adodbRecordsetSeek(rs *adodbRecordset, args []Value) {
	if rs == nil || rs.state != adStateOpen || len(args) == 0 {
		return
	}
	if len(args) == 1 && args[0].Type == VTInteger {
		_ = vm.dispatchADODBRecordsetPropertySet(rs, "AbsolutePosition", args[0])
		return
	}
	fieldName := strings.TrimSpace(rs.index)
	if fieldName == "" && len(rs.columns) > 0 {
		fieldName = rs.columns[0]
	}
	if idx := strings.Index(fieldName, ","); idx >= 0 {
		fieldName = fieldName[:idx]
	}
	fieldName = strings.ToLower(strings.Trim(fieldName, " []`\""))
	if fieldName == "" {
		return
	}
	needle := strings.TrimSpace(vm.valueToString(args[0]))
	for i := 0; i < len(rs.data); i++ {
		row := rs.data[i]
		if row == nil {
			continue
		}
		if strings.EqualFold(strings.TrimSpace(row[fieldName].String()), needle) {
			rs.currentRow = i
			rs.bookmark = i + 1
			rs.eof = false
			rs.bof = false
			return
		}
	}
	rs.currentRow = rs.recordCount
	rs.eof = true
	rs.bof = false
}

// adodbRecordsetSupports returns true for the in-memory features the VM exposes.
func (vm *VM) adodbRecordsetSupports(rs *adodbRecordset, args []Value) bool {
	if rs == nil || len(args) == 0 {
		return false
	}
	return rs.state == adStateOpen || len(rs.columns) > 0
}

// adodbRecordsetUpdateBatch finalizes batched edit state for disconnected recordsets.
func (vm *VM) adodbRecordsetUpdateBatch(rs *adodbRecordset) {
	if rs == nil {
		return
	}
	rs.editMode = adEditNone
	rs.status = adRecOK
}

// adodbRecordsetPersistXML serializes one recordset snapshot to a compact XML format.
func (vm *VM) adodbRecordsetPersistXML(rs *adodbRecordset) string {
	var builder strings.Builder
	builder.Grow(256 + len(rs.data)*64)
	builder.WriteString("<?xml version=\"1.0\" encoding=\"utf-8\"?><recordset>")
	for i := 0; i < len(rs.data); i++ {
		row := rs.data[i]
		if row == nil {
			continue
		}
		builder.WriteString("<row>")
		for c := 0; c < len(rs.columns); c++ {
			columnName := rs.columns[c]
			lowerName := strings.ToLower(columnName)
			builder.WriteString("<field name=\"")
			builder.WriteString(adodbXMLEscape(columnName))
			builder.WriteString("\">")
			builder.WriteString(adodbXMLEscape(row[lowerName].String()))
			builder.WriteString("</field>")
		}
		builder.WriteString("</row>")
	}
	builder.WriteString("</recordset>")
	return builder.String()
}

// adodbResolvePersistPath resolves a Save target path using the current host sandbox when available.
func (vm *VM) adodbResolvePersistPath(path string) (string, bool) {
	trimmed := strings.TrimSpace(path)
	if trimmed == "" {
		return "", false
	}
	if filepath.IsAbs(trimmed) {
		return trimmed, true
	}
	if vm.host != nil && vm.host.Server() != nil {
		return vm.host.Server().MapPath(trimmed), true
	}
	absolutePath, err := filepath.Abs(trimmed)
	if err != nil {
		return "", false
	}
	return absolutePath, true
}

// adodbParseConnectionStringParams parses one semicolon-delimited connection string.
func adodbParseConnectionStringParams(connStr string) map[string]string {
	params := make(map[string]string)
	parts := strings.Split(connStr, ";")
	for _, part := range parts {
		if idx := strings.Index(part, "="); idx > 0 {
			key := strings.ToLower(strings.TrimSpace(part[:idx]))
			val := strings.Trim(strings.TrimSpace(part[idx+1:]), "{}")
			params[key] = val
		}
	}
	return params
}

// adodbApplyPlatformAccessProvider rewrites Access provider strings according to
// global.adodb_platform_architecture (auto/386/amd64) on Windows runtime.
func adodbApplyPlatformAccessProvider(connStr string) string {
	params := adodbParseConnectionStringParams(connStr)
	provider := strings.TrimSpace(params["provider"])
	lowerProvider := strings.ToLower(provider)
	if !strings.Contains(lowerProvider, "microsoft.jet.oledb") && !strings.Contains(lowerProvider, "microsoft.ace.oledb") {
		return connStr
	}

	effectiveArch := adodbEffectiveAccessArchitecture()
	targetProvider := "Microsoft.ACE.OLEDB.12.0"
	if effectiveArch == "386" {
		targetProvider = "Microsoft.Jet.OLEDB.4.0"
	}
	if strings.EqualFold(provider, targetProvider) {
		return connStr
	}
	return adodbUpsertConnectionStringValue(connStr, "Provider", targetProvider)
}

// adodbEffectiveAccessArchitecture resolves the effective Access provider architecture.
// Values: "386" forces Jet, "amd64" forces ACE. "auto" follows runtime bitness.
func adodbEffectiveAccessArchitecture() string {
	configured := adodbConfiguredPlatformArchitecture()
	if configured == "386" || configured == "amd64" {
		return configured
	}
	if runtime.GOARCH == "386" {
		return "386"
	}
	return "amd64"
}

// adodbConfiguredPlatformArchitecture returns normalized config mode from
// global.adodb_platform_architecture.
func adodbConfiguredPlatformArchitecture() string {
	if strings.TrimSpace(adodbPlatformArchitectureTestOverride) != "" {
		override := strings.ToLower(strings.TrimSpace(adodbPlatformArchitectureTestOverride))
		switch override {
		case "386", "amd64", "auto":
			return override
		default:
			return "auto"
		}
	}
	adodbPlatformArchitectureOnce.Do(func() {
		mode := "auto"
		if raw, ok := loadAxConfigValue("global.adodb_platform_architecture"); ok {
			mode = strings.ToLower(strings.TrimSpace(fmt.Sprintf("%v", raw)))
		}
		switch mode {
		case "386", "amd64", "auto":
			adodbPlatformArchitectureCached = mode
		default:
			adodbPlatformArchitectureCached = "auto"
		}
	})
	return adodbPlatformArchitectureCached
}

// adodbNormalizeAccessConnectionString applies the Access sharing options required for updateable cursors.
func adodbNormalizeAccessConnectionString(connStr string) string {
	updated := adodbUpsertConnectionStringValue(connStr, "Mode", "Share Deny None")
	updated = adodbUpsertConnectionStringValue(updated, "Jet OLEDB:Database Locking Mode", "1")
	return updated
}

// adodbConnectionDebugSummary returns a compact provider/data-source summary for error context.
func adodbConnectionDebugSummary(connStr string) string {
	params := adodbParseConnectionStringParams(connStr)
	provider := strings.TrimSpace(params["provider"])
	if provider == "" {
		provider = "<unknown>"
	}
	dataSource := strings.TrimSpace(params["data source"])
	if dataSource == "" {
		dataSource = strings.TrimSpace(params["datasource"])
	}
	if dataSource == "" {
		dataSource = "<unknown>"
	}
	return "provider=" + provider + "; data source=" + dataSource
}

// adodbAlternateAccessProviderConnectionString swaps ACE/JET providers when an
// Access OLE connection fails so the runtime can retry with the alternate provider.
func adodbAlternateAccessProviderConnectionString(connStr string) (string, bool) {
	params := adodbParseConnectionStringParams(connStr)
	provider := strings.TrimSpace(params["provider"])
	switch {
	case strings.EqualFold(provider, "Microsoft.ACE.OLEDB.12.0"):
		return adodbUpsertConnectionStringValue(connStr, "Provider", "Microsoft.Jet.OLEDB.4.0"), true
	case strings.EqualFold(provider, "Microsoft.Jet.OLEDB.4.0"):
		return adodbUpsertConnectionStringValue(connStr, "Provider", "Microsoft.ACE.OLEDB.12.0"), true
	default:
		return "", false
	}
}

// adodbUpsertConnectionStringValue replaces or appends one semicolon-delimited connection-string key.
func adodbUpsertConnectionStringValue(connStr string, key string, value string) string {
	parts := strings.Split(connStr, ";")
	for i := 0; i < len(parts); i++ {
		segment := strings.TrimSpace(parts[i])
		if segment == "" {
			continue
		}
		idx := strings.Index(segment, "=")
		if idx <= 0 {
			continue
		}
		if strings.EqualFold(strings.TrimSpace(segment[:idx]), key) {
			parts[i] = key + "=" + value
			return strings.Join(parts, ";")
		}
	}
	trimmed := strings.TrimSpace(connStr)
	if trimmed == "" {
		return key + "=" + value
	}
	if strings.HasSuffix(trimmed, ";") {
		return connStr + key + "=" + value
	}
	return connStr + ";" + key + "=" + value
}

// adodbParseConnectionInteger reads the first parseable integer from the provided connection keys.
func adodbParseConnectionInteger(params map[string]string, keys ...string) int {
	for _, key := range keys {
		if raw, ok := params[key]; ok {
			parsed, err := strconv.Atoi(strings.TrimSpace(raw))
			if err == nil {
				return parsed
			}
		}
	}
	return 0
}

// adodbXMLEscape escapes XML-sensitive characters for Recordset.Save persistence.
func adodbXMLEscape(value string) string {
	value = strings.ReplaceAll(value, "&", "&amp;")
	value = strings.ReplaceAll(value, "<", "&lt;")
	value = strings.ReplaceAll(value, ">", "&gt;")
	value = strings.ReplaceAll(value, `"`, "&quot;")
	return value
}

// Access/OLE specific implementation
func (vm *VM) adodbOpenAccessDatabase(conn *adodbConnection, connStr string) {
	connStr = adodbNormalizeAccessConnectionString(connStr)
	conn.connectionString = connStr
	err := vm.ensureCOMRequestThread()
	if err != nil {
		vm.raise(vbscript.InternalError,
			"COM initialization failed: "+err.Error())
		return
	}

	unknown, err := oleutil.CreateObject("ADODB.Connection")
	if err != nil {
		vm.adodbConnectionClose(conn)
		vm.raise(vbscript.InternalError,
			"ADODB.Connection OLE creation failed: "+err.Error())
		return
	}
	dispatch, err := unknown.QueryInterface(ole.IID_IDispatch)
	unknown.Release()
	if err != nil {
		vm.adodbConnectionClose(conn)
		vm.raise(vbscript.InternalError,
			"ADODB.Connection OLE interface failed: "+err.Error())
		return
	}

	modeRes, modeErr := oleutil.PutProperty(dispatch, "Mode", int32(adModeShareDenyNone))
	if modeRes != nil {
		modeRes.Clear()
	}
	if modeErr != nil {
		dispatch.Release()
		vm.adodbConnectionClose(conn)
		vm.raise(vbscript.InternalError,
			"ADODB.Connection OLE Mode failed: "+modeErr.Error())
		return
	}

	// Do NOT set CursorLocation = adUseClient on the Connection.
	// Activating the Microsoft Cursor Service (MSDAER.dll) at the Connection level
	// causes Connection.Execute to go through the Cursor Service, which is an
	// STA COM object. Calling it from an MTA-initialised thread causes an
	// intra-process apartment marshaling deadlock for complex queries (e.g. LEFT
	// JOINs with ORDER BY): the Cursor Service's STA thread waits for a message
	// the locked OS thread can never deliver.  We use adUseServer (the default)
	// so that ACE OLEDB handles the rowset in-process with no STA involvement.
	// Rows are then eagerly materialised into rs.data by adodbPopulateRecordsetFromOLE.
	_, err = oleutil.CallMethod(dispatch, "Open", connStr)
	if err != nil {
		if fallbackConnStr, ok := adodbAlternateAccessProviderConnectionString(connStr); ok {
			_, fallbackErr := oleutil.CallMethod(dispatch, "Open", fallbackConnStr)
			if fallbackErr == nil {
				conn.connectionString = fallbackConnStr
				conn.oleConnection = dispatch
				conn.state = adStateOpen
				return
			}
		}
		dispatch.Release()
		vm.adodbConnectionClose(conn)
		summary := adodbConnectionDebugSummary(connStr)
		vm.raise(vbscript.InternalError,
			"ADODB.Connection OLE Open failed: "+err.Error()+" ("+summary+")")
		return
	}

	conn.oleConnection = dispatch
	conn.state = adStateOpen
}

// adodbPopulateRecordsetFromOLE reads column names and all rows from an OLE ADO Recordset
// into the VM's native cache so the recordset can be served without further OLE calls.
func (vm *VM) adodbPopulateRecordsetFromOLE(rs *adodbRecordset) {
	if rs.oleRecordset == nil {
		return
	}

	// Fetch the Fields collection.
	fieldsRes, _ := oleutil.GetProperty(rs.oleRecordset, "Fields")
	if fieldsRes == nil {
		return
	}
	fields := fieldsRes.ToIDispatch()
	fieldsRes.Clear()
	if fields == nil {
		return
	}
	defer fields.Release()

	// Read column count. VARIANT.Value() properly coerces the OLE type (VT_I4 → int32).
	count := 0
	countRes, _ := oleutil.GetProperty(fields, "Count")
	if countRes != nil {
		switch v := countRes.Value().(type) {
		case int32:
			count = int(v)
		case int64:
			count = int(v)
		default:
			count = int(countRes.Val)
		}
		countRes.Clear()
	}

	// Populate column name slice.
	rs.columns = make([]string, count)
	for i := 0; i < count; i++ {
		itemRes, _ := oleutil.GetProperty(fields, "Item", i)
		if itemRes == nil {
			continue
		}
		item := itemRes.ToIDispatch()
		itemRes.Clear()
		if item == nil {
			continue
		}
		nameRes, _ := oleutil.GetProperty(item, "Name")
		if nameRes != nil {
			name := strings.TrimSpace(nameRes.ToString())
			if name == "" {
				switch nv := nameRes.Value().(type) {
				case string:
					name = strings.TrimSpace(nv)
				case []uint16:
					name = strings.TrimSpace(string(utf16.Decode(nv)))
				}
			}
			if name == "" {
				name = "field" + strconv.Itoa(i)
			}
			rs.columns[i] = name
			nameRes.Clear()
		}
		item.Release()
	}
	vm.adodbRecordsetRebuildColumnIndex(rs)

	// Read GetRows from a clone cursor so the original cursor remains intact for
	// compatibility fallback when SAFEARRAY decoding is unavailable.
	var cloneRecordset *ole.IDispatch
	cloneRes, cloneErr := oleutil.CallMethod(rs.oleRecordset, "Clone")
	if cloneErr == nil && cloneRes != nil {
		cloneDisp := cloneRes.ToIDispatch()
		if cloneDisp != nil {
			cloneDisp.AddRef()
			cloneRecordset = cloneDisp
		}
		cloneRes.Clear()
	}

	if cloneRecordset != nil {
		// Fast path: fetch all rows in a single COM boundary crossing.
		rowsRes, rowsErr := oleutil.CallMethod(cloneRecordset, "GetRows", int32(-1))
		if rowsErr == nil && rowsRes != nil {
			ok := vm.adodbPopulateRecordsetFromOLEGetRowsResult(rs, rowsRes, count)
			rowsRes.Clear()
			if ok {
				cloneRecordset.Release()
				vm.adodbFinalizeOLEMaterializedRecordset(rs)
				return
			}
		}
		cloneRecordset.Release()
	}

	vm.adodbPopulateRecordsetFromOLEFieldWalk(rs, fields, count)
	vm.adodbFinalizeOLEMaterializedRecordset(rs)
}

// adodbPopulateRecordsetFromOLEGetRowsResult converts one OLE GetRows VARIANT payload
// into the in-memory row cache. It expects field-major ordering: [field][row].
func (vm *VM) adodbPopulateRecordsetFromOLEGetRowsResult(rs *adodbRecordset, rowsRes *ole.VARIANT, count int) bool {
	if rs == nil || rowsRes == nil || count < 0 {
		return false
	}
	arr := rowsRes.ToArray()
	if arr == nil {
		return false
	}

	if count == 0 {
		rs.data = rs.data[:0]
		return true
	}

	dimsPtr, err := arr.GetDimensions()
	if err != nil || dimsPtr == nil {
		return false
	}
	if *dimsPtr == 0 {
		rs.data = rs.data[:0]
		return true
	}

	fieldsLen, err := arr.TotalElements(1)
	if err != nil || fieldsLen < 0 {
		return false
	}
	if fieldsLen == 0 {
		rs.data = rs.data[:0]
		return true
	}
	if int(fieldsLen) != count {
		return false
	}

	if decodedValues, decodedRows, ok := adodbDecodeGetRowsVariant(rowsRes, count); ok {
		if decodedRows <= 0 {
			rs.data = rs.data[:0]
			return true
		}
		return vm.adodbHydrateRecordsetDataFromFieldMajorValues(rs, decodedValues, count, decodedRows)
	}

	values := arr.ToValueArray()
	if len(values) == 0 {
		rs.data = rs.data[:0]
		return true
	}

	rowCount := len(values) / count
	if *dimsPtr >= 2 {
		rowsLen, err := arr.TotalElements(2)
		if err == nil && rowsLen >= 0 {
			rowCount = int(rowsLen)
		}
	}
	if rowCount <= 0 {
		rs.data = rs.data[:0]
		return true
	}
	if len(values) < rowCount*count {
		// go-ole currently exposes SafeArray conversion helpers tuned for 1D arrays.
		// If a 2D payload cannot be read reliably, preserve legacy provider behavior.
		return false
	}

	return vm.adodbHydrateRecordsetDataFromFieldMajorValues(rs, values, count, rowCount)
}

// adodbHydrateRecordsetDataFromFieldMajorValues materialises rows from a field-major
// flattened value list where index = field + row*fieldCount.
func (vm *VM) adodbHydrateRecordsetDataFromFieldMajorValues(rs *adodbRecordset, values []interface{}, fieldCount int, rowCount int) bool {
	if rs == nil || fieldCount < 0 || rowCount < 0 {
		return false
	}
	if fieldCount == 0 {
		rs.data = rs.data[:0]
		return true
	}
	need := fieldCount * rowCount
	if need < 0 || len(values) < need {
		return false
	}

	data := make([]map[string]Value, rowCount)
	for rowIdx := 0; rowIdx < rowCount; rowIdx++ {
		row := make(map[string]Value, fieldCount)
		base := rowIdx * fieldCount
		for colIdx := 0; colIdx < fieldCount; colIdx++ {
			key := strings.ToLower(strings.TrimSpace(rs.columns[colIdx]))
			if key == "" {
				continue
			}
			row[key] = vm.adodbValueToVMValue(values[base+colIdx])
		}
		data[rowIdx] = row
	}
	rs.data = data
	return true
}

// adodbPopulateRecordsetFromOLEFieldWalk preserves the compatibility fallback that
// walks fields and values row-by-row when GetRows cannot be decoded safely.
func (vm *VM) adodbPopulateRecordsetFromOLEFieldWalk(rs *adodbRecordset, fields *ole.IDispatch, count int) {
	// adodbOLEMaxRows caps the fetch loop as a safety guard: if the OLE provider
	// returns a broken cursor where EOF never transitions to true (e.g. due to a
	// nil/VT_NULL variant), the loop terminates rather than running forever.
	const adodbOLEMaxRows = 500_000
	rs.data = make([]map[string]Value, 0)
	eofRes, _ := oleutil.GetProperty(rs.oleRecordset, "EOF")
	if eofRes == nil {
		return
	}
	for rowIdx := 0; rowIdx < adodbOLEMaxRows; rowIdx++ {
		eofRaw := eofRes.Value()
		if eofRaw == nil || vm.asBool(vm.adodbValueToVMValue(eofRaw)) {
			break
		}
		row := make(map[string]Value, count)
		for i := 0; i < count; i++ {
			itemRes, _ := oleutil.GetProperty(fields, "Item", i)
			if itemRes == nil {
				continue
			}
			item := itemRes.ToIDispatch()
			itemRes.Clear()
			if item == nil {
				continue
			}
			valRes, _ := oleutil.GetProperty(item, "Value")
			if valRes != nil {
				key := strings.ToLower(strings.TrimSpace(rs.columns[i]))
				if key != "" {
					row[key] = vm.adodbValueToVMValue(valRes.Value())
				}
				valRes.Clear()
			}
			item.Release()
		}
		rs.data = append(rs.data, row)
		_, _ = oleutil.CallMethod(rs.oleRecordset, "MoveNext")
		eofRes.Clear()
		eofRes, _ = oleutil.GetProperty(rs.oleRecordset, "EOF")
		if eofRes == nil {
			break
		}
	}
	if eofRes != nil {
		eofRes.Clear()
	}
}

// adodbFinalizeOLEMaterializedRecordset updates EOF/BOF and cursor positioning
// after eager OLE materialization into native VM storage.
func (vm *VM) adodbFinalizeOLEMaterializedRecordset(rs *adodbRecordset) {
	rs.recordCount = len(rs.data)
	if rs.recordCount > 0 {
		rs.currentRow = 0
		rs.eof = false
		rs.bof = false
	} else {
		rs.currentRow = -1
		rs.eof = true
		rs.bof = true
	}
}

// comInitialize initialises COM on the current OS thread using the Single-Threaded
// Apartment (STA) model.
//
// ADODB.Connection has ThreadingModel=Apartment, meaning it must live in an STA.
// When the calling thread is MTA, COM creates a separate host-STA thread for the
// Connection object and marshals every method call across apartments.  During a
// complex query (e.g. LEFT JOIN + ORDER BY on a large Access database) ACE fires
// progress/notification callbacks back into the Connection object.  If the
// host-STA thread is busy executing Execute(), it cannot process those incoming
// callbacks, causing a permanent deadlock.
//
// With STA, ADODB.Connection is created directly in the calling thread's apartment
// (no marshaling proxy).  ACE runs inline on the same thread; any callbacks from
// ACE to Connection are nested in-apartment calls on the same stack — COM handles
// these without a message pump.  This is exactly the model used by IIS Classic ASP
// (one STA worker thread per request).
//
// The earlier MoveNext-in-a-loop deadlock that motivated the MTA switch no longer
// applies: all rows are now eagerly materialised in adodbPopulateRecordsetFromOLE
// on the OS-locked goroutine thread before the IDispatch is released.
func comInitialize() (bool, error) {
	err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	if err != nil {
		// RPC_E_CHANGED_MODE (0x80010106): this OS thread was already initialised
		// under a different apartment model.  COM is still usable; skip the
		// corresponding CoUninitialize.
		if strings.Contains(err.Error(), "0x80010106") {
			return false, nil
		}
		// 0x8001010E: call made from the wrong apartment — non-fatal, carry on.
		if strings.Contains(err.Error(), "0x8001010E") {
			return false, nil
		}
		return false, err
	}
	return true, nil
}
