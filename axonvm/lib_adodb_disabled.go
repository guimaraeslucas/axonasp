//go:build wasm || lib_adodb_disabled

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

	"g3pix.com.br/axonasp/vbscript"
)

// adodbError is the disabled stub for an ADODB.Error runtime instance.
type adodbError struct{}

// adodbConnection is the disabled stub for an ADODB.Connection runtime instance.
// Only fields required by always-compiled code paths are retained.
type adodbConnection struct {
	errors []adodbError
}

// adodbParameter is the disabled stub for an ADODB.Parameter runtime instance.
type adodbParameter struct{}

// adodbRecordset is the disabled stub for an ADODB.Recordset runtime instance.
// The columns field is retained so builtins.go can enumerate it safely.
type adodbRecordset struct {
	columns []string
}

// adodbCommand is the disabled stub for an ADODB.Command runtime instance.
// The parameters field is retained so builtins.go can enumerate it safely.
type adodbCommand struct {
	parameters []*adodbParameter
}

// adodbFieldProxy is the disabled stub for an ADODB Field proxy runtime instance.
type adodbFieldProxy struct{}

// newADODBConnection panics with ErrLibraryDisabled because ADODB support is compiled out.
func (vm *VM) newADODBConnection() Value {
	panic(&VMError{
		Code:        vbscript.ActiveXCannotCreateObject,
		Number:      int(ErrLibraryDisabled),
		Description: fmt.Sprintf(ErrLibraryDisabled.String(), "adodb"),
		Source:      "ADODB.Connection",
	})
}

// newADODBOLEConnection panics with ErrLibraryDisabled because ADODB support is compiled out.
func (vm *VM) newADODBOLEConnection() Value {
	panic(&VMError{
		Code:        vbscript.ActiveXCannotCreateObject,
		Number:      int(ErrLibraryDisabled),
		Description: fmt.Sprintf(ErrLibraryDisabled.String(), "adodb"),
		Source:      "ADODB.Connection (OLE)",
	})
}

// newADODBRecordset panics with ErrLibraryDisabled because ADODB support is compiled out.
func (vm *VM) newADODBRecordset() Value {
	panic(&VMError{
		Code:        vbscript.ActiveXCannotCreateObject,
		Number:      int(ErrLibraryDisabled),
		Description: fmt.Sprintf(ErrLibraryDisabled.String(), "adodb"),
		Source:      "ADODB.Recordset",
	})
}

// newADODBCommand panics with ErrLibraryDisabled because ADODB support is compiled out.
func (vm *VM) newADODBCommand() Value {
	panic(&VMError{
		Code:        vbscript.ActiveXCannotCreateObject,
		Number:      int(ErrLibraryDisabled),
		Description: fmt.Sprintf(ErrLibraryDisabled.String(), "adodb"),
		Source:      "ADODB.Command",
	})
}

// newADODBFieldProxy returns an empty value because ADODB support is compiled out.
func (vm *VM) newADODBFieldProxy(rs *adodbRecordset, name string) Value {
	return Value{Type: VTEmpty}
}

// adodbConnectionOpen is a no-op stub because ADODB support is compiled out.
func (vm *VM) adodbConnectionOpen(conn *adodbConnection) {}

// adodbConnectionClose is a no-op stub because ADODB support is compiled out.
func (vm *VM) adodbConnectionClose(conn *adodbConnection) {}

// dispatchADODBMethod returns unhandled for all calls because ADODB support is compiled out.
func (vm *VM) dispatchADODBMethod(objID int64, member string, args []Value) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

// dispatchADODBPropertyGet returns unhandled for all property reads because ADODB support is compiled out.
func (vm *VM) dispatchADODBPropertyGet(objID int64, member string) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

// dispatchADODBPropertySet returns unhandled for all property writes because ADODB support is compiled out.
func (vm *VM) dispatchADODBPropertySet(objID int64, member string, val Value) bool {
	return false
}

// dispatchADODBErrorsCollectionMethod returns unhandled because ADODB support is compiled out.
func (vm *VM) dispatchADODBErrorsCollectionMethod(objID int64, member string, args []Value) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

// dispatchADODBErrorsCollectionPropertyGet returns unhandled because ADODB support is compiled out.
func (vm *VM) dispatchADODBErrorsCollectionPropertyGet(objID int64, member string) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

// dispatchADODBErrorPropertyGet returns unhandled because ADODB support is compiled out.
func (vm *VM) dispatchADODBErrorPropertyGet(objID int64, member string) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

// dispatchADODBFieldsCollectionMethod returns unhandled because ADODB support is compiled out.
func (vm *VM) dispatchADODBFieldsCollectionMethod(objID int64, member string, args []Value) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

// dispatchADODBParametersCollectionMethod returns unhandled because ADODB support is compiled out.
func (vm *VM) dispatchADODBParametersCollectionMethod(objID int64, member string, args []Value) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

// dispatchADODBFieldsCollectionPropertyGet returns unhandled because ADODB support is compiled out.
func (vm *VM) dispatchADODBFieldsCollectionPropertyGet(objID int64, member string) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

// dispatchADODBParametersCollectionPropertyGet returns unhandled because ADODB support is compiled out.
func (vm *VM) dispatchADODBParametersCollectionPropertyGet(objID int64, member string) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

// dispatchADODBFieldMethod returns unhandled because ADODB support is compiled out.
func (vm *VM) dispatchADODBFieldMethod(objID int64, member string, args []Value) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

// dispatchADODBFieldPropertyGet returns unhandled because ADODB support is compiled out.
func (vm *VM) dispatchADODBFieldPropertyGet(objID int64, member string) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

// dispatchADODBFieldPropertySet returns unhandled because ADODB support is compiled out.
func (vm *VM) dispatchADODBFieldPropertySet(objID int64, member string, val Value) bool {
	return false
}

// CleanupRequestResources releases non-ADODB request-scoped native resources.
// ADODB cleanup is skipped because the library is compiled out.
func (vm *VM) CleanupRequestResources() {
	if vm == nil {
		return
	}
	vm.cleanupG3ImageResources()
	vm.cleanupG3ZSTDResources()
	clear(vm.nativeObjectProxies)
	vm.nextDynamicNativeID = 20000
}

// ensureCOMRequestThread is a no-op because ADODB (and COM) support is compiled out.
func (vm *VM) ensureCOMRequestThread() error {
	return nil
}

// releaseCOMRequestThread is a no-op because ADODB (and COM) support is compiled out.
func (vm *VM) releaseCOMRequestThread() {
}

func (vm *VM) startSTAWorker() {}
func (vm *VM) stopSTAWorker()  {}
func (vm *VM) runOnSTA(f func()) {
	f()
}
