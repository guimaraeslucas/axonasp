/*
 * AxonASP Server - Version 1.0
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
	"fmt"
	"strings"
	"sync"
)

// DictionaryLibrary wraps Dictionary for ASPLibrary interface compatibility
type DictionaryLibrary struct {
	dict *Dictionary
}

// NewDictionary creates a new Dictionary object (wrapped as ASPLibrary)
func NewDictionary(ctx *ExecutionContext) *DictionaryLibrary {
	return &DictionaryLibrary{
		dict: &Dictionary{
			store:       make(map[string]interface{}),
			order:       make([]string, 0),
			compareMode: 0, // Binary mode by default
		},
	}
}

// CallMethod calls a method on the Dictionary
func (dl *DictionaryLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	result := dl.dict.CallMethod(name, args...)
	return result, nil
}

// GetProperty gets a property from the Dictionary
func (dl *DictionaryLibrary) GetProperty(name string) interface{} {
	return dl.dict.GetProperty(name)
}

// SetProperty sets a property on the Dictionary
func (dl *DictionaryLibrary) SetProperty(name string, value interface{}) error {
	dl.dict.SetProperty(name, value)
	return nil
}

// Enumeration returns all keys for For Each support
func (dl *DictionaryLibrary) Enumeration() []interface{} {
	return dl.dict.Enumeration()
}

// Dictionary implements the Scripting.Dictionary COM object
type Dictionary struct {
	store       map[string]interface{}
	mutex       sync.RWMutex
	compareMode int // 0=Binary, 1=TextCompare
	order       []string
}

// GetProperty gets a property value from the Dictionary
func (d *Dictionary) GetProperty(name string) interface{} {
	d.mutex.RLock()
	defer d.mutex.RUnlock()

	lowerName := strings.ToLower(name)

	switch lowerName {
	case "count":
		return int64(len(d.store))
	case "comparemode":
		return int64(d.compareMode)
	default:
		return nil
	}
}

// SetProperty sets a property value on the Dictionary
func (d *Dictionary) SetProperty(name string, value interface{}) {
	d.mutex.Lock()
	defer d.mutex.Unlock()

	lowerName := strings.ToLower(name)

	switch lowerName {
	case "comparemode":
		switch v := value.(type) {
		case int64:
			d.compareMode = int(v)
		case int:
			d.compareMode = v
		case float64:
			d.compareMode = int(v)
		}
	}
}

// CallMethod calls a method on the Dictionary
func (d *Dictionary) CallMethod(name string, args ...interface{}) interface{} {
	lowerName := strings.ToLower(name)

	switch lowerName {
	case "add":
		return d.Add(args)
	case "exists":
		return d.Exists(args)
	case "item", "": // Default indexing: dict("key") is equivalent to dict.Item("key")
		return d.Item(args)
	case "remove":
		return d.Remove(args)
	case "removeall":
		return d.RemoveAll(args)
	case "keys":
		return d.Keys(args)
	case "items":
		return d.Items(args)
	default:
		return nil
	}
}

// Add adds a key-value pair to the Dictionary
func (d *Dictionary) Add(args []interface{}) interface{} {
	if len(args) < 2 {
		return nil
	}

	d.mutex.Lock()
	defer d.mutex.Unlock()

	key := d.keyToString(args[0])
	if _, exists := d.store[key]; !exists {
		d.order = append(d.order, key)
	}
	d.store[key] = args[1]
	return nil
}

// Exists checks if a key exists in the Dictionary
func (d *Dictionary) Exists(args []interface{}) interface{} {
	if len(args) < 1 {
		return false
	}

	d.mutex.RLock()
	defer d.mutex.RUnlock()

	key := d.keyToString(args[0])
	_, exists := d.store[key]
	return exists
}

// Item gets or sets an item in the Dictionary
func (d *Dictionary) Item(args []interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}

	d.mutex.RLock()
	defer d.mutex.RUnlock()

	key := d.keyToString(args[0])
	value, exists := d.store[key]
	if !exists {
		return nil
	}
	return value
}

// Remove removes a key from the Dictionary
func (d *Dictionary) Remove(args []interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}

	d.mutex.Lock()
	defer d.mutex.Unlock()

	key := d.keyToString(args[0])
	delete(d.store, key)
	d.removeKeyOrder(key)
	return nil
}

// RemoveAll clears all items from the Dictionary
func (d *Dictionary) RemoveAll(args []interface{}) interface{} {
	d.mutex.Lock()
	defer d.mutex.Unlock()

	d.store = make(map[string]interface{})
	d.order = d.order[:0]
	return nil
}

// Keys returns all keys from the Dictionary
func (d *Dictionary) Keys(args []interface{}) interface{} {
	d.mutex.RLock()
	defer d.mutex.RUnlock()

	keys := make([]interface{}, 0, len(d.order))
	for _, k := range d.order {
		keys = append(keys, k)
	}
	return keys
}

// Items returns all values from the Dictionary
func (d *Dictionary) Items(args []interface{}) interface{} {
	d.mutex.RLock()
	defer d.mutex.RUnlock()

	items := make([]interface{}, 0, len(d.store))
	for _, v := range d.store {
		items = append(items, v)
	}
	return items
}

// keyToString converts a key to string using case-insensitive comparison if needed
func (d *Dictionary) keyToString(key interface{}) string {
	switch v := key.(type) {
	case string:
		if d.compareMode == 1 { // TextCompare
			return strings.ToLower(v)
		}
		return v
	case int64:
		return fmt.Sprintf("%d", v)
	case int:
		return fmt.Sprintf("%d", v)
	case float64:
		return fmt.Sprintf("%v", v)
	default:
		keyStr := fmt.Sprintf("%v", v)
		if d.compareMode == 1 { // TextCompare
			return strings.ToLower(keyStr)
		}
		return keyStr
	}
}

// Enumeration returns all keys for For Each support
func (d *Dictionary) Enumeration() []interface{} {
	d.mutex.RLock()
	defer d.mutex.RUnlock()

	keys := make([]interface{}, 0, len(d.order))
	for _, k := range d.order {
		keys = append(keys, k)
	}
	return keys
}

func (d *Dictionary) removeKeyOrder(key string) {
	for i, k := range d.order {
		if k == key {
			d.order = append(d.order[:i], d.order[i+1:]...)
			return
		}
	}
}
