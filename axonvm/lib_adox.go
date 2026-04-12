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
	"fmt"
	"runtime"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

const adSchemaTables = 20

type ADOXCatalog struct {
	vm               *VM
	activeConnection Value // Either a string or a VTNativeObject
	tables           Value
}

// newADOXCatalogObject instantiates the ADOX.Catalog custom functions library.
func (vm *VM) newADOXCatalogObject() Value {
	obj := &ADOXCatalog{
		vm:               vm,
		activeConnection: Value{Type: VTEmpty},
		tables:           Value{Type: VTEmpty},
	}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.adoxCatalogItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet acts as a getter.
func (c *ADOXCatalog) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(strings.TrimSpace(propertyName)) {
	case "tables":
		return c.getTables()
	case "activeconnection":
		return c.activeConnection
	}
	return NewEmpty()
}

// DispatchPropertySet acts as a setter.
func (c *ADOXCatalog) DispatchPropertySet(propertyName string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	switch strings.ToLower(strings.TrimSpace(propertyName)) {
	case "activeconnection":
		c.activeConnection = args[0]
		c.tables = Value{Type: VTEmpty}
		return true
	}
	return false
}

// DispatchMethod handles methods for Catalog.
func (c *ADOXCatalog) DispatchMethod(methodName string, args []Value) Value {
	return NewEmpty()
}

func (c *ADOXCatalog) getTables() Value {
	if c.tables.Type != VTEmpty {
		return c.tables
	}

	tables := c.loadTables()
	c.tables = c.vm.newADOXTablesObject(tables)
	return c.tables
}

func (c *ADOXCatalog) loadTables() []*ADOXTable {
	if runtime.GOOS == "windows" {
		oleConn, cleanup := c.resolveOLEConnection()
		if cleanup != nil {
			defer cleanup()
		}
		if oleConn != nil {
			return listADOXTables(oleConn)
		}
	}
	return c.listTablesNative()
}

func (c *ADOXCatalog) listTablesNative() []*ADOXTable {
	if c.activeConnection.Type != VTNativeObject {
		return []*ADOXTable{}
	}

	conn, exists := c.vm.adodbConnectionItems[c.activeConnection.Num]
	if !exists || conn == nil || conn.db == nil {
		return []*ADOXTable{}
	}

	if strings.Contains(strings.ToLower(conn.dbDriver), "sqlite") {
		rows, err := conn.db.Query("SELECT name, type FROM sqlite_master WHERE type='table' OR type='view'")
		if err != nil {
			return []*ADOXTable{}
		}
		defer rows.Close()

		tables := make([]*ADOXTable, 0)
		for rows.Next() {
			var name, typeStr string
			if err := rows.Scan(&name, &typeStr); err == nil {
				tType := "TABLE"
				if strings.ToLower(typeStr) == "view" {
					tType = "VIEW"
				}
				tables = append(tables, &ADOXTable{Name: name, Type: tType})
			}
		}
		return tables
	}

	return []*ADOXTable{}
}

func (c *ADOXCatalog) resolveOLEConnection() (*ole.IDispatch, func()) {
	switch c.activeConnection.Type {
	case VTString:
		connStr := strings.TrimSpace(c.activeConnection.String())
		if connStr != "" {
			return c.openTemporaryOLEConnection(connStr)
		}
	case VTNativeObject:
		conn, exists := c.vm.adodbConnectionItems[c.activeConnection.Num]
		if exists && conn != nil {
			if conn.oleConnection != nil {
				return conn.oleConnection, nil
			}
			if strings.TrimSpace(conn.connectionString) != "" {
				return c.openTemporaryOLEConnection(conn.connectionString)
			}
		}
	}
	return nil, nil
}

func (c *ADOXCatalog) openTemporaryOLEConnection(connStr string) (*ole.IDispatch, func()) {
	// Create an ADODB.Connection dynamically
	tempConnVal := c.vm.newADODBConnection()
	tempConn, ok := c.vm.adodbConnectionItems[tempConnVal.Num]
	if !ok || tempConn == nil {
		return nil, nil
	}

	tempConn.connectionString = connStr
	c.vm.adodbConnectionOpen(tempConn)
	if tempConn.oleConnection == nil {
		c.vm.adodbConnectionClose(tempConn)
		return nil, nil
	}

	cleanup := func() {
		c.vm.adodbConnectionClose(tempConn)
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
	defer result.Clear()

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
		isEOF := false
		if eofResult != nil {
			isEOF = variantToBool(eofResult)
			eofResult.Clear()
		}
		if isEOF {
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

		moveNextRes, _ := oleutil.CallMethod(rsDisp, "MoveNext")
		if moveNextRes != nil {
			moveNextRes.Clear()
		}
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
	defer fieldsResult.Clear()

	fieldsDisp := fieldsResult.ToIDispatch()
	if fieldsDisp == nil {
		return nil
	}
	defer fieldsDisp.Release()

	fieldResult, err := oleutil.GetProperty(fieldsDisp, "Item", fieldName)
	if err != nil {
		return nil
	}
	defer fieldResult.Clear()

	fieldDisp := fieldResult.ToIDispatch()
	if fieldDisp == nil {
		return nil
	}
	defer fieldDisp.Release()

	valueResult, err := oleutil.GetProperty(fieldDisp, "Value")
	if err != nil {
		return nil
	}
	defer valueResult.Clear()

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
	vm    *VM
	items []*ADOXTable
}

func (vm *VM) newADOXTablesObject(items []*ADOXTable) Value {
	obj := &ADOXTables{
		vm:    vm,
		items: items,
	}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.adoxTablesItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

func (t *ADOXTables) DispatchPropertyGet(name string) Value {
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "count":
		return NewInteger(int64(len(t.items)))
	case "item":
		return NewEmpty()
	}
	return t.DispatchMethod(name, nil)
}

func (t *ADOXTables) DispatchMethod(name string, args []Value) Value {
	method := strings.ToLower(strings.TrimSpace(name))
	if method == "" || method == "item" {
		if len(args) < 1 {
			return NewEmpty()
		}

		val := args[0]
		switch val.Type {
		case VTInteger:
			idx := int(val.Num)
			if idx >= 0 && idx < len(t.items) {
				return t.vm.newADOXTableObject(t.items[idx])
			}
		case VTString:
			key := strings.ToLower(strings.TrimSpace(val.String()))
			for _, item := range t.items {
				if strings.ToLower(item.Name) == key {
					return t.vm.newADOXTableObject(item)
				}
			}
		}
		return NewEmpty()
	}

	return NewEmpty()
}

// ADOXTable represents a single ADOX table item.
type ADOXTable struct {
	Name string
	Type string
}

type ADOXTableWrapper struct {
	vm   *VM
	item *ADOXTable
}

func (vm *VM) newADOXTableObject(item *ADOXTable) Value {
	obj := &ADOXTableWrapper{
		vm:   vm,
		item: item,
	}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.adoxTableItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

func (t *ADOXTableWrapper) DispatchPropertyGet(name string) Value {
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "name":
		return NewString(t.item.Name)
	case "type":
		return NewString(t.item.Type)
	}
	return NewEmpty()
}

func (t *ADOXTableWrapper) DispatchMethod(name string, args []Value) Value {
	return NewEmpty()
}
