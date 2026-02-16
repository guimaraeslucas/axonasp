/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimaraes - G3pix Ltda
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
	"fmt"
	"runtime"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

const adSchemaTables = 20

// ADOXCatalog implements a minimal ADOX.Catalog compatible object.
type ADOXCatalog struct {
	ctx              *ExecutionContext
	activeConnection interface{}
	tables           *ADOXTables
}

// NewADOXCatalog creates a new ADOX catalog instance.
func NewADOXCatalog(ctx *ExecutionContext) *ADOXCatalog {
	return &ADOXCatalog{ctx: ctx}
}

// GetProperty returns catalog properties (Tables, ActiveConnection).
func (c *ADOXCatalog) GetProperty(name string) interface{} {
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "tables":
		return c.getTables()
	case "activeconnection":
		return c.activeConnection
	}
	return nil
}

// SetProperty sets catalog properties (ActiveConnection).
func (c *ADOXCatalog) SetProperty(name string, value interface{}) error {
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "activeconnection":
		c.activeConnection = value
		c.tables = nil
	}
	return nil
}

// CallMethod supports ADOX catalog methods (none required for FileSearch).
func (c *ADOXCatalog) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return nil, nil
}

func (c *ADOXCatalog) getTables() *ADOXTables {
	if c.tables != nil {
		return c.tables
	}

	tables := c.loadTables()
	c.tables = &ADOXTables{items: tables}
	return c.tables
}

func (c *ADOXCatalog) loadTables() []*ADOXTable {
	if runtime.GOOS != "windows" {
		return []*ADOXTable{}
	}

	oleConn, cleanup := c.resolveOLEConnection()
	if cleanup != nil {
		defer cleanup()
	}
	if oleConn == nil {
		return []*ADOXTable{}
	}

	tables := listADOXTables(oleConn)
	return tables
}

func (c *ADOXCatalog) resolveOLEConnection() (*ole.IDispatch, func()) {
	if c.activeConnection == nil {
		return nil, nil
	}

	switch conn := c.activeConnection.(type) {
	case *ADODBConnection:
		if conn.oleConnection != nil {
			return conn.oleConnection, nil
		}
		if strings.TrimSpace(conn.ConnectionString) != "" {
			return openTemporaryOLEConnection(conn.ConnectionString)
		}
	case *ADOConnection:
		if conn.lib != nil && conn.lib.oleConnection != nil {
			return conn.lib.oleConnection, nil
		}
		if conn.lib != nil && strings.TrimSpace(conn.lib.ConnectionString) != "" {
			return openTemporaryOLEConnection(conn.lib.ConnectionString)
		}
	case string:
		if strings.TrimSpace(conn) != "" {
			return openTemporaryOLEConnection(conn)
		}
	case *ole.IDispatch:
		return conn, nil
	}

	return nil, nil
}

func openTemporaryOLEConnection(connStr string) (*ole.IDispatch, func()) {
	tempConn := NewADODBConnection(nil)
	tempConn.ConnectionString = connStr
	tempConn.openDatabase()
	if tempConn.oleConnection == nil {
		tempConn.CallMethod("close")
		return nil, nil
	}

	cleanup := func() {
		tempConn.CallMethod("close")
	}
	return tempConn.oleConnection, cleanup
}

func listADOXTables(oleConn *ole.IDispatch) []*ADOXTable {
	if oleConn == nil {
		return []*ADOXTable{}
	}

	result, err := oleutil.CallMethod(oleConn, "OpenSchema", int32(adSchemaTables))
	if err != nil {
		return []*ADOXTable{}
	}

	rsDisp := result.ToIDispatch()
	if rsDisp == nil {
		return []*ADOXTable{}
	}
	defer rsDisp.Release()

	tables := make([]*ADOXTable, 0)
	for {
		eofResult, eofErr := oleutil.GetProperty(rsDisp, "EOF")
		if eofErr != nil {
			break
		}
		if variantToBool(eofResult) {
			break
		}

		nameVal := oleRecordsetFieldValue(rsDisp, "TABLE_NAME")
		typeVal := oleRecordsetFieldValue(rsDisp, "TABLE_TYPE")
		name := strings.TrimSpace(fmt.Sprintf("%v", nameVal))
		if name != "" {
			tableType := strings.ToUpper(strings.TrimSpace(fmt.Sprintf("%v", typeVal)))
			if tableType == "" {
				tableType = "TABLE"
			}
			tables = append(tables, &ADOXTable{Name: name, Type: tableType})
		}

		_, _ = oleutil.CallMethod(rsDisp, "MoveNext")
	}

	return tables
}

func oleRecordsetFieldValue(rs *ole.IDispatch, fieldName string) interface{} {
	if rs == nil {
		return nil
	}

	fieldsResult, err := oleutil.GetProperty(rs, "Fields")
	if err != nil {
		return nil
	}
	fieldsDisp := fieldsResult.ToIDispatch()
	if fieldsDisp == nil {
		return nil
	}
	defer fieldsDisp.Release()

	fieldResult, err := oleutil.GetProperty(fieldsDisp, "Item", fieldName)
	if err != nil {
		return nil
	}
	fieldDisp := fieldResult.ToIDispatch()
	if fieldDisp == nil {
		return nil
	}
	defer fieldDisp.Release()

	valueResult, err := oleutil.GetProperty(fieldDisp, "Value")
	if err != nil {
		return nil
	}
	return valueResult.Value()
}

func variantToBool(value *ole.VARIANT) bool {
	if value == nil {
		return true
	}

	switch v := value.Value().(type) {
	case bool:
		return v
	case int:
		return v != 0
	case int32:
		return v != 0
	case int64:
		return v != 0
	case float32:
		return v != 0
	case float64:
		return v != 0
	case string:
		return strings.TrimSpace(v) != "0" && v != ""
	default:
		return value.Val != 0
	}
}

// ADOXTables represents the catalog tables collection.
type ADOXTables struct {
	items []*ADOXTable
}

func (t *ADOXTables) GetProperty(name string) interface{} {
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "count":
		return len(t.items)
	case "item":
		return nil
	}
	return nil
}

func (t *ADOXTables) SetProperty(name string, value interface{}) error {
	return nil
}

func (t *ADOXTables) CallMethod(name string, args ...interface{}) (interface{}, error) {
	method := strings.ToLower(strings.TrimSpace(name))
	if method == "" {
		method = "item"
	}

	switch method {
	case "item":
		if len(args) < 1 {
			return nil, nil
		}
		switch idx := args[0].(type) {
		case int:
			if idx >= 0 && idx < len(t.items) {
				return t.items[idx], nil
			}
		case int32:
			if int(idx) >= 0 && int(idx) < len(t.items) {
				return t.items[int(idx)], nil
			}
		case int64:
			if int(idx) >= 0 && int(idx) < len(t.items) {
				return t.items[int(idx)], nil
			}
		case float64:
			if int(idx) >= 0 && int(idx) < len(t.items) {
				return t.items[int(idx)], nil
			}
		default:
			key := strings.ToLower(strings.TrimSpace(fmt.Sprintf("%v", args[0])))
			for _, item := range t.items {
				if strings.ToLower(item.Name) == key {
					return item, nil
				}
			}
		}
		return nil, nil
	case "count":
		return len(t.items), nil
	}

	return nil, nil
}

func (t *ADOXTables) Enumeration() []interface{} {
	items := make([]interface{}, 0, len(t.items))
	for _, item := range t.items {
		items = append(items, item)
	}
	return items
}

// ADOXTable represents a single ADOX table item.
type ADOXTable struct {
	Name string
	Type string
}

func (t *ADOXTable) GetProperty(name string) interface{} {
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "name":
		return t.Name
	case "type":
		return t.Type
	}
	return nil
}

func (t *ADOXTable) SetProperty(name string, value interface{}) error {
	return nil
}

func (t *ADOXTable) CallMethod(name string, args ...interface{}) (interface{}, error) {
	if strings.TrimSpace(name) == "" {
		return nil, nil
	}
	return nil, nil
}
