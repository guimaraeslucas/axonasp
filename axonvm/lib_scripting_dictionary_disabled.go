//go:build lib_scripting_dictionary_disabled

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

import "strings"

// scriptingDictionary is the disabled placeholder for Scripting.Dictionary.
type scriptingDictionary struct {
	keys        []Value
	values      []Value
	index       map[string]int
	compareMode int
}

func newScriptingDictionary() *scriptingDictionary {
	return &scriptingDictionary{index: make(map[string]int)}
}

func (d *scriptingDictionary) normalizeKey(key string) string {
	if d != nil && d.compareMode == 1 {
		return key
	}
	return strings.ToLower(key)
}

func (d *scriptingDictionary) keyIndex(key string) int {
	if d == nil || d.index == nil {
		return -1
	}
	if pos, ok := d.index[d.normalizeKey(key)]; ok {
		return pos
	}
	return -1
}

func (d *scriptingDictionary) itemGet(vm *VM, key Value) Value {
	if d == nil {
		return Value{Type: VTEmpty}
	}
	pos := d.keyIndex(key.String())
	if pos >= 0 && pos < len(d.values) {
		return d.values[pos]
	}
	return Value{Type: VTEmpty}
}

func (d *scriptingDictionary) itemSet(vm *VM, key Value, val Value) {
	if d == nil {
		return
	}
	pos := d.keyIndex(key.String())
	if pos >= 0 && pos < len(d.values) {
		d.values[pos] = val
		return
	}
	d.addEntry(vm, key, val)
}

func (d *scriptingDictionary) addEntry(vm *VM, key Value, val Value) {
	if d == nil {
		return
	}
	if d.index == nil {
		d.index = make(map[string]int)
	}
	pos := len(d.keys)
	d.keys = append(d.keys, key)
	d.values = append(d.values, val)
	d.index[d.normalizeKey(key.String())] = pos
}

func (d *scriptingDictionary) rebuildIndex() {
	if d == nil {
		return
	}
	if d.index == nil {
		d.index = make(map[string]int)
	}
	clear(d.index)
	for i, k := range d.keys {
		d.index[d.normalizeKey(k.String())] = i
	}
}

// newDictionaryObject fails because Scripting.Dictionary is disabled at compile time.
func (vm *VM) newDictionaryObject() Value {
	panicLibraryDisabled("scripting_dictionary", "Scripting.Dictionary")
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchDictionaryMethod(objID int64, member string, args []Value) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

func (vm *VM) dispatchDictionaryPropertyGet(objID int64, member string) (Value, bool) {
	return Value{Type: VTEmpty}, false
}

func (vm *VM) dispatchDictionaryPropertySet(objID int64, member string, val Value) bool {
	return false
}
