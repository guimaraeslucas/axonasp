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
	"strings"

	"g3pix.com.br/axonasp/vbscript"
)

// scriptingDictionary is the runtime state for one Scripting.Dictionary instance.
// Keys are stored in insertion order. The index map holds the normalized key
// (lowercased when TextCompare) mapped to the position in the keys/values slices.
// This gives O(1) Exists/Item access and preserves Keys()/Items() ordering.
type scriptingDictionary struct {
	keys        []Value
	values      []Value
	index       map[string]int
	compareMode int // 0 = BinaryCompare (case-sensitive), 1 = TextCompare (case-insensitive)
}

// newScriptingDictionary allocates a fresh, empty Scripting.Dictionary instance.
func newScriptingDictionary() *scriptingDictionary {
	return &scriptingDictionary{
		keys:   make([]Value, 0, 8),
		values: make([]Value, 0, 8),
		index:  make(map[string]int, 8),
	}
}

// normalizeKey applies the current CompareMode to produce the lookup key.
func (d *scriptingDictionary) normalizeKey(key string) string {
	if d.compareMode == 1 {
		return strings.ToLower(key)
	}
	return key
}

// keyIndex returns the slice position for the given key, or -1 if not found.
func (d *scriptingDictionary) keyIndex(key string) int {
	if pos, ok := d.index[d.normalizeKey(key)]; ok {
		return pos
	}
	return -1
}

// itemGet returns the value for key. If the key does not exist, it is auto-created
// with Empty value (classic VBScript Dictionary late-binding behavior).
func (d *scriptingDictionary) itemGet(key Value) Value {
	pos := d.keyIndex(key.String())
	if pos >= 0 {
		return d.values[pos]
	}
	// Auto-create with Empty (VBScript behavior: reading an absent key creates it).
	d.addEntry(key, Value{Type: VTEmpty})
	return Value{Type: VTEmpty}
}

// itemSet assigns value to key, creating the entry if it does not exist.
func (d *scriptingDictionary) itemSet(key Value, val Value) {
	pos := d.keyIndex(key.String())
	if pos >= 0 {
		d.values[pos] = val
		return
	}
	d.addEntry(key, val)
}

// addEntry appends a new key-value pair. Caller must ensure the key does not exist.
func (d *scriptingDictionary) addEntry(key Value, val Value) {
	pos := len(d.keys)
	d.keys = append(d.keys, key)
	d.values = append(d.values, val)
	d.index[d.normalizeKey(key.String())] = pos
}

// rebuildIndex reconstructs the index map from the current keys slice.
// Called after Remove or CompareMode change.
func (d *scriptingDictionary) rebuildIndex() {
	d.index = make(map[string]int, len(d.keys))
	for i, k := range d.keys {
		d.index[d.normalizeKey(k.String())] = i
	}
}

// newDictionaryObject stores a fresh Dictionary instance in the VM and returns its handle.
func (vm *VM) newDictionaryObject() Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.dictionaryItems[objID] = newScriptingDictionary()
	return Value{Type: VTNativeObject, Num: objID}
}

// dispatchDictionaryMethod handles all method calls and the default property
// (member == "") for Scripting.Dictionary instances identified by objID.
// Returns (result, true) when objID belongs to a Dictionary; (Empty, false) otherwise.
func (vm *VM) dispatchDictionaryMethod(objID int64, member string, args []Value) (Value, bool) {
	d, exists := vm.dictionaryItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}

	// Default property: dict(key) get or dict(key) = value set (from OpArraySet).
	// OpArraySet passes args as [indexes..., value], so len==2 means set; len==1 means get.
	if member == "" {
		if len(args) == 1 {
			return d.itemGet(args[0]), true
		}
		if len(args) >= 2 {
			// Last arg is the value (OpArraySet convention); first arg is the key.
			d.itemSet(args[0], args[len(args)-1])
			return Value{Type: VTEmpty}, true
		}
		return Value{Type: VTEmpty}, true
	}

	switch {
	// Add(key, value) — adds a new entry; error if key already exists.
	case strings.EqualFold(member, "Add"):
		if len(args) < 2 {
			vm.raise(vbscript.WrongNumberOfParameters, "Dictionary.Add requires 2 arguments: key, value")
			return Value{Type: VTEmpty}, true
		}
		if d.keyIndex(args[0].String()) >= 0 {
			vm.raise(vbscript.ThisKeyAlreadyAssociatedWithAnElement, "This key is already associated with an element of this collection")
			return Value{Type: VTEmpty}, true
		}
		d.addEntry(args[0], args[1])
		return Value{Type: VTEmpty}, true

	// Exists(key) — returns True if key is present.
	case strings.EqualFold(member, "Exists"):
		if len(args) < 1 {
			vm.raise(vbscript.WrongNumberOfParameters, "Dictionary.Exists requires 1 argument: key")
			return Value{Type: VTEmpty}, true
		}
		return NewBool(d.keyIndex(args[0].String()) >= 0), true

	// Remove(key) — removes key; error if not found.
	case strings.EqualFold(member, "Remove"):
		if len(args) < 1 {
			vm.raise(vbscript.WrongNumberOfParameters, "Dictionary.Remove requires 1 argument: key")
			return Value{Type: VTEmpty}, true
		}
		pos := d.keyIndex(args[0].String())
		if pos < 0 {
			vm.raise(vbscript.ElementWasNotFound, "The specified key was not found in the Dictionary")
			return Value{Type: VTEmpty}, true
		}
		d.keys = append(d.keys[:pos], d.keys[pos+1:]...)
		d.values = append(d.values[:pos], d.values[pos+1:]...)
		d.rebuildIndex()
		return Value{Type: VTEmpty}, true

	// RemoveAll() — clears all entries.
	case strings.EqualFold(member, "RemoveAll"):
		d.keys = d.keys[:0]
		d.values = d.values[:0]
		d.index = make(map[string]int, 8)
		return Value{Type: VTEmpty}, true

	// Keys() — returns a zero-indexed Variant array of all keys.
	case strings.EqualFold(member, "Keys"):
		count := len(d.keys)
		arr := NewVBArrayFromValues(0, make([]Value, count))
		copy(arr.Values, d.keys)
		return Value{Type: VTArray, Arr: arr}, true

	// Items() — returns a zero-indexed Variant array of all values.
	case strings.EqualFold(member, "Items"):
		count := len(d.values)
		arr := NewVBArrayFromValues(0, make([]Value, count))
		copy(arr.Values, d.values)
		return Value{Type: VTArray, Arr: arr}, true

	// Item(key) — get or set the item for key.
	case strings.EqualFold(member, "Item"):
		if len(args) == 1 {
			// If it's a simple get, we return the value.
			// But if this is part of a "dict.Item(key) = val" assignment,
			// the compiler might have used OpCallMember and then OpMemberSet.
			// However, in VBScript "dict.Item(key)" is usually a GET.
			// If we want to support "dict.Item(key) = val", we need a proxy.
			// Let's return the actual value for 1 arg, BUT for Key we definitely need a proxy if 1 arg.
			return d.itemGet(args[0]), true
		}
		if len(args) >= 2 {
			// OpArraySet convention: last arg is value, first arg is key.
			d.itemSet(args[0], args[len(args)-1])
			return Value{Type: VTEmpty}, true
		}
		vm.raise(vbscript.WrongNumberOfParameters, "Dictionary.Item requires 1 argument: key")
		return Value{Type: VTEmpty}, true

	// Key(oldKey) = newKey — rename an existing key.
	case strings.EqualFold(member, "Key"):
		if len(args) == 1 {
			// Proxy for "dict.Key(old) = new"
			return vm.newNativeObjectProxy(objID, "Key", args), true
		}
		if len(args) >= 2 {
			oldKey := args[0].String()
			newKey := args[1].String()
			pos := d.keyIndex(oldKey)
			if pos < 0 {
				vm.raise(vbscript.ElementWasNotFound, "The specified key was not found in the Dictionary")
				return Value{Type: VTEmpty}, true
			}
			if d.keyIndex(newKey) >= 0 {
				vm.raise(vbscript.ThisKeyAlreadyAssociatedWithAnElement, "The new key is already associated with an element of this collection")
				return Value{Type: VTEmpty}, true
			}
			// Update key while preserving value at the same position.
			d.keys[pos] = args[1]
			d.rebuildIndex()
			return Value{Type: VTEmpty}, true
		}
		vm.raise(vbscript.WrongNumberOfParameters, "Dictionary.Key requires 2 arguments: oldKey, newKey")
		return Value{Type: VTEmpty}, true

	// Count — returns number of entries.
	case strings.EqualFold(member, "Count"):
		return NewInteger(int64(len(d.keys))), true

	// CompareMode — read property via method call (no-arg).
	case strings.EqualFold(member, "CompareMode"):
		return NewInteger(int64(d.compareMode)), true
	}

	vm.raise(vbscript.ObjectDoesntSupportThisPropertyOrMethod, "Scripting.Dictionary does not support '"+member+"'")
	return Value{Type: VTEmpty}, true
}

// dispatchDictionaryPropertyGet handles property reads (OpMemberGet) for Dictionary.
// Returns (result, true) when objID is a Dictionary.
func (vm *VM) dispatchDictionaryPropertyGet(objID int64, member string) (Value, bool) {
	d, exists := vm.dictionaryItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}

	switch {
	case strings.EqualFold(member, "Count"):
		return NewInteger(int64(len(d.keys))), true
	case strings.EqualFold(member, "CompareMode"):
		return NewInteger(int64(d.compareMode)), true
	case strings.EqualFold(member, "Keys"):
		count := len(d.keys)
		arr := NewVBArrayFromValues(0, make([]Value, count))
		copy(arr.Values, d.keys)
		return Value{Type: VTArray, Arr: arr}, true
	case strings.EqualFold(member, "Items"):
		count := len(d.values)
		arr := NewVBArrayFromValues(0, make([]Value, count))
		copy(arr.Values, d.values)
		return Value{Type: VTArray, Arr: arr}, true
	case strings.EqualFold(member, "Item"):
		return vm.newNativeObjectProxy(objID, "Item", nil), true
	case strings.EqualFold(member, "Key"):
		return vm.newNativeObjectProxy(objID, "Key", nil), true
	case member == "__default__" || member == "":
		// Internal VM sentinel used by OpCoerceToValue to probe for a default scalar value.
		// Dictionary has no scalar default — return Empty without auto-creating a key.
		return Value{Type: VTEmpty}, true
	}

	// Default: treat as Item(key) access — e.g. dict.Key reads Item(key).
	// Per VBScript spec, any unknown member on a Dictionary is treated as Item.
	return d.itemGet(NewString(member)), true
}

// dispatchDictionaryPropertySet handles property Let assignments (OpMemberSet) for Dictionary.
// Returns true when objID is a Dictionary.
func (vm *VM) dispatchDictionaryPropertySet(objID int64, member string, val Value) bool {
	d, exists := vm.dictionaryItems[objID]
	if !exists {
		return false
	}

	switch {
	// Item(key) = value — set or create entry via member assignment.
	case strings.EqualFold(member, "Item"):
		// The key is not known from this calling convention; skip (use Add/direct assign instead).
		return true

	// CompareMode = n — change comparison mode; must be set before adding items.
	case strings.EqualFold(member, "CompareMode"):
		if len(d.keys) > 0 {
			vm.raise(vbscript.InvalidProcedureCallOrArgument, "The CompareMode property cannot be changed if the Dictionary object already contains data.")
			return true
		}
		mode := int(val.Num)
		if mode != d.compareMode {
			d.compareMode = mode
			d.rebuildIndex()
		}
		return true

	// Key(oldKey) = newKey — rename an existing key.
	case strings.EqualFold(member, "Key"):
		// The old key is embedded in the member name after "Key"; not standard—
		// handled via method dispatch instead.
		return true
	}

	// Default: treat as Item(member) = val — dictionary member-name assignment.
	d.itemSet(NewString(member), val)
	return true
}
