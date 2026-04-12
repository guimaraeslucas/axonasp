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
package asp

import (
	"strings"
	"sync"
)

// ApplicationValueType identifies the stored primitive variant kind.
type ApplicationValueType byte

const (
	// ApplicationValueEmpty stores an empty value.
	ApplicationValueEmpty ApplicationValueType = iota
	// ApplicationValueString stores a string.
	ApplicationValueString
	// ApplicationValueInteger stores an integer.
	ApplicationValueInteger
	// ApplicationValueDouble stores a floating-point value.
	ApplicationValueDouble
	// ApplicationValueBool stores a boolean value.
	ApplicationValueBool
)

// ApplicationValue stores a VM-compatible primitive without using interface allocations.
// ApplicationValueArray stores a VBScript array (including multi-dimensional) recursively.
const ApplicationValueArray ApplicationValueType = 5

type ApplicationValue struct {
	Type     ApplicationValueType
	Num      int64
	Flt      float64
	Str      string
	Arr      []ApplicationValue // Elements when Type == ApplicationValueArray.
	ArrLower int                // Lower bound of this array dimension.
}

// NewApplicationArray creates an ApplicationValue holding a VBScript array dimension.
func NewApplicationArray(lower int, arr []ApplicationValue) ApplicationValue {
	return ApplicationValue{Type: ApplicationValueArray, ArrLower: lower, Arr: arr}
}

// NewApplicationEmpty creates an empty application value.
func NewApplicationEmpty() ApplicationValue {
	return ApplicationValue{Type: ApplicationValueEmpty}
}

// NewApplicationString creates a string application value.
func NewApplicationString(v string) ApplicationValue {
	return ApplicationValue{Type: ApplicationValueString, Str: v}
}

// NewApplicationInteger creates an integer application value.
func NewApplicationInteger(v int64) ApplicationValue {
	return ApplicationValue{Type: ApplicationValueInteger, Num: v}
}

// NewApplicationDouble creates a floating-point application value.
func NewApplicationDouble(v float64) ApplicationValue {
	return ApplicationValue{Type: ApplicationValueDouble, Flt: v}
}

// NewApplicationBool creates a boolean application value.
func NewApplicationBool(v bool) ApplicationValue {
	if v {
		return ApplicationValue{Type: ApplicationValueBool, Num: 1}
	}
	return ApplicationValue{Type: ApplicationValueBool, Num: 0}
}

// Bool returns the boolean representation for boolean values.
func (v ApplicationValue) Bool() bool {
	return v.Num != 0
}

// normalizeApplicationKey normalizes a key for case-insensitive lookup semantics.
func normalizeApplicationKey(key string) string {
	return strings.ToLower(key)
}

// Application maintains shared state between all users.
type Application struct {
	// Contents stores standard application variables.
	contents map[string]ApplicationValue
	// StaticObjects stores values declared as static object entries.
	staticObjects map[string]ApplicationValue
	// mutex protects contents and staticObjects.
	mutex sync.RWMutex
	// stateMu protects lock state bookkeeping.
	stateMu sync.Mutex
	// stateCond blocks other requests while one request owns the Application lock.
	stateCond *sync.Cond
	// lockOwner identifies the request-level Server that currently owns the lock.
	lockOwner *Server
	// locked indicates whether at least one lock is active.
	locked bool
	// lockCount tracks nested Lock/Unlock pairs.
	lockCount int
}

// NewApplication creates a VM-native Application object.
func NewApplication() *Application {
	app := &Application{
		contents:      make(map[string]ApplicationValue),
		staticObjects: make(map[string]ApplicationValue),
	}
	app.stateCond = sync.NewCond(&app.stateMu)
	return app
}

// Lock enters an application-wide critical section and supports nested locks.
func (app *Application) Lock() {
	app.LockForServer(nil)
}

// LockForServer enters an application-wide critical section owned by one request server.
func (app *Application) LockForServer(owner *Server) {
	app.stateMu.Lock()
	defer app.stateMu.Unlock()
	for app.lockCount > 0 && app.lockOwner != owner {
		app.stateCond.Wait()
	}

	app.lockOwner = owner
	app.lockCount++
	app.locked = app.lockCount > 0
}

// Unlock leaves an application-wide critical section for one nesting level.
func (app *Application) Unlock() {
	app.UnlockForServer(nil)
}

// UnlockForServer leaves one application-wide critical section nesting level for one request server.
func (app *Application) UnlockForServer(owner *Server) {
	app.stateMu.Lock()
	defer app.stateMu.Unlock()
	if app.lockCount == 0 {
		return
	}
	if app.lockOwner != nil && owner != nil && app.lockOwner != owner {
		return
	}

	app.lockCount--
	app.locked = app.lockCount > 0
	if app.lockCount == 0 {
		app.lockOwner = nil
		app.stateCond.Broadcast()
	}
}

// WaitForServer blocks until the current Application lock is released or owned by the same request server.
func (app *Application) WaitForServer(owner *Server) {
	app.stateMu.Lock()
	defer app.stateMu.Unlock()
	for app.lockCount > 0 && app.lockOwner != owner {
		app.stateCond.Wait()
	}
}

// IsLocked reports whether the Application object is currently locked.
func (app *Application) IsLocked() bool {
	app.stateMu.Lock()
	defer app.stateMu.Unlock()

	return app.locked
}

// GetLockCount returns the current nested lock count.
func (app *Application) GetLockCount() int {
	app.stateMu.Lock()
	defer app.stateMu.Unlock()

	return app.lockCount
}

// Set stores a value in Application.Contents using case-insensitive keys.
func (app *Application) Set(key string, value ApplicationValue) {
	app.mutex.Lock()
	defer app.mutex.Unlock()

	app.contents[normalizeApplicationKey(key)] = value
}

// Get returns a value from Application.Contents using case-insensitive keys.
func (app *Application) Get(key string) (ApplicationValue, bool) {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	value, ok := app.contents[normalizeApplicationKey(key)]
	return value, ok
}

// ContainsContent reports whether a Contents key exists.
func (app *Application) ContainsContent(key string) bool {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	_, ok := app.contents[normalizeApplicationKey(key)]
	return ok
}

// Remove deletes one key from Application.Contents.
func (app *Application) Remove(key string) {
	app.mutex.Lock()
	defer app.mutex.Unlock()

	delete(app.contents, normalizeApplicationKey(key))
}

// RemoveAll clears all keys from Application.Contents.
func (app *Application) RemoveAll() {
	app.mutex.Lock()
	defer app.mutex.Unlock()

	app.contents = make(map[string]ApplicationValue)
}

// Count returns total number of entries in Contents and StaticObjects.
func (app *Application) Count() int {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	return len(app.contents) + len(app.staticObjects)
}

// AddStaticObject inserts a value into Application.StaticObjects.
func (app *Application) AddStaticObject(key string, value ApplicationValue) {
	app.mutex.Lock()
	defer app.mutex.Unlock()

	app.staticObjects[normalizeApplicationKey(key)] = value
}

// GetStaticObject returns a value from Application.StaticObjects.
func (app *Application) GetStaticObject(key string) (ApplicationValue, bool) {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	value, ok := app.staticObjects[normalizeApplicationKey(key)]
	return value, ok
}

// ContainsStaticObject reports whether a StaticObjects key exists.
func (app *Application) ContainsStaticObject(key string) bool {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	_, ok := app.staticObjects[normalizeApplicationKey(key)]
	return ok
}

// GetContentsCopy returns a snapshot copy of Contents for safe enumeration.
func (app *Application) GetContentsCopy() map[string]ApplicationValue {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	copyMap := make(map[string]ApplicationValue, len(app.contents))
	for key, value := range app.contents {
		copyMap[key] = value
	}
	return copyMap
}

// GetStaticObjectsCopy returns a snapshot copy of StaticObjects for safe enumeration.
func (app *Application) GetStaticObjectsCopy() map[string]ApplicationValue {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	copyMap := make(map[string]ApplicationValue, len(app.staticObjects))
	for key, value := range app.staticObjects {
		copyMap[key] = value
	}
	return copyMap
}

// GetAllKeys returns all keys from both Contents and StaticObjects.
func (app *Application) GetAllKeys() []string {
	app.mutex.RLock()
	defer app.mutex.RUnlock()

	keys := make([]string, 0, len(app.contents)+len(app.staticObjects))
	for key := range app.contents {
		keys = append(keys, key)
	}
	for key := range app.staticObjects {
		keys = append(keys, key)
	}
	return keys
}
