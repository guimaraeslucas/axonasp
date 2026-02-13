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
	"strings"
	"sync"
)

// normalizeKey converts a key to lowercase for case-insensitive storage
func normalizeKey(key string) string {
	return strings.ToLower(key)
}

// ApplicationObject represents the ASP Application object
// Provides application-wide state with Lock/Unlock for thread safety
type ApplicationObject struct {
	// Contents collection (standard variables)
	contents map[string]interface{}
	// StaticObjects collection (objects declared in global.asa with <OBJECT>)
	staticObjects map[string]interface{}
	// Mutex for Lock/Unlock methods
	mutex sync.RWMutex
	// Separate mutex for lock-state bookkeeping
	stateMu sync.Mutex
	// Lock state tracking
	locked bool
	// Lock count to support nested locks
	lockCount int
}

// NewApplicationObject creates a new Application object
func NewApplicationObject() *ApplicationObject {
	return &ApplicationObject{
		contents:      make(map[string]interface{}),
		staticObjects: make(map[string]interface{}),
		locked:        false,
	}
}

// Lock prevents other clients from modifying Application variables
// Supports nested locks - each Lock must have a corresponding Unlock
func (app *ApplicationObject) Lock() {
	app.stateMu.Lock()
	defer app.stateMu.Unlock()
	app.lockCount++
	app.locked = app.lockCount > 0
}

// Unlock allows other clients to modify Application variables
// Must be called the same number of times as Lock was called
func (app *ApplicationObject) Unlock() {
	app.stateMu.Lock()
	defer app.stateMu.Unlock()
	if app.lockCount > 0 {
		app.lockCount--
	}
	app.locked = app.lockCount > 0
}

// Get retrieves a value from Application.Contents (case-insensitive)
func (app *ApplicationObject) Get(key string) interface{} {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	// Normalize key to lowercase for case-insensitive lookup
	keyLower := normalizeKey(key)
	return app.contents[keyLower]
}

// Set stores a value in Application.Contents (case-insensitive)
func (app *ApplicationObject) Set(key string, value interface{}) {
	// Use RLock for read check, then upgrade if needed
	app.mutex.Lock()
	defer app.mutex.Unlock()

	// Normalize key to lowercase for case-insensitive storage
	keyLower := normalizeKey(key)
	app.contents[keyLower] = value
}

// Remove removes a key from Application.Contents
func (app *ApplicationObject) Remove(key string) {
	app.mutex.Lock()
	defer app.mutex.Unlock()

	keyLower := normalizeKey(key)
	delete(app.contents, keyLower)
}

// RemoveAll clears all variables from Application.Contents
func (app *ApplicationObject) RemoveAll() {
	app.mutex.Lock()
	defer app.mutex.Unlock()

	app.contents = make(map[string]interface{})
}

// GetContents returns the Contents collection for enumeration
// Returns a copy to avoid concurrent modification issues
func (app *ApplicationObject) GetContents() map[string]interface{} {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	// Return a copy of contents
	contentsCopy := make(map[string]interface{})
	for k, v := range app.contents {
		contentsCopy[k] = v
	}
	return contentsCopy
}

// GetStaticObjects returns the StaticObjects collection for enumeration
// For backwards compatibility, returns all keys (Contents + StaticObjects)
// Returns a copy to avoid concurrent modification issues
func (app *ApplicationObject) GetStaticObjects() map[string]interface{} {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	// Return combined map of both contents and static objects
	combined := make(map[string]interface{})
	for k, v := range app.contents {
		combined[k] = v
	}
	for k, v := range app.staticObjects {
		combined[k] = v
	}
	return combined
}

// AddStaticObject adds an object to the StaticObjects collection
// This is typically called when parsing global.asa <OBJECT> tags
func (app *ApplicationObject) AddStaticObject(key string, obj interface{}) {
	app.mutex.Lock()
	defer app.mutex.Unlock()

	keyLower := normalizeKey(key)
	app.staticObjects[keyLower] = obj
}

// GetAllKeys returns all keys from both Contents and StaticObjects
// Used for variable enumeration (For Each)
func (app *ApplicationObject) GetAllKeys() []string {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	keys := make([]string, 0, len(app.contents)+len(app.staticObjects))

	// Add contents keys
	for k := range app.contents {
		keys = append(keys, k)
	}

	// Add static objects keys
	for k := range app.staticObjects {
		keys = append(keys, k)
	}

	return keys
}

// Count returns the total number of items in Contents and StaticObjects
func (app *ApplicationObject) Count() int {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	return len(app.contents) + len(app.staticObjects)
}

// IsLocked returns whether the Application object is currently locked
func (app *ApplicationObject) IsLocked() bool {
	app.stateMu.Lock()
	defer app.stateMu.Unlock()
	return app.locked
}

// GetLockCount returns the current lock count
func (app *ApplicationObject) GetLockCount() int {
	app.stateMu.Lock()
	defer app.stateMu.Unlock()
	return app.lockCount
}

// GetContentsCopy returns a deep copy of the Contents collection
// Safe for external iteration without locks
func (app *ApplicationObject) GetContentsCopy() map[string]interface{} {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	copy := make(map[string]interface{})
	for k, v := range app.contents {
		copy[k] = v
	}
	return copy
}

// GetStaticObjectsCopy returns a copy of only the StaticObjects collection
// Safe for external iteration without locks
func (app *ApplicationObject) GetStaticObjectsCopy() map[string]interface{} {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	copy := make(map[string]interface{})
	for k, v := range app.staticObjects {
		copy[k] = v
	}
	return copy
}

// GetContentKeys returns only the keys from Contents collection
func (app *ApplicationObject) GetContentKeys() []string {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	keys := make([]string, 0, len(app.contents))
	for k := range app.contents {
		keys = append(keys, k)
	}
	return keys
}

// GetStaticObjectKeys returns only the keys from StaticObjects collection
func (app *ApplicationObject) GetStaticObjectKeys() []string {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	keys := make([]string, 0, len(app.staticObjects))
	for k := range app.staticObjects {
		keys = append(keys, k)
	}
	return keys
}

// ContainsContent checks if a key exists in Contents collection
func (app *ApplicationObject) ContainsContent(key string) bool {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	keyLower := normalizeKey(key)
	_, exists := app.contents[keyLower]
	return exists
}

// ContainsStaticObject checks if a key exists in StaticObjects collection
func (app *ApplicationObject) ContainsStaticObject(key string) bool {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	keyLower := normalizeKey(key)
	_, exists := app.staticObjects[keyLower]
	return exists
}
