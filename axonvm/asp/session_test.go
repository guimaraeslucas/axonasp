/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
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
	"os"
	"path/filepath"
	"testing"
	"time"
)

// TestSessionCaseInsensitiveContents verifies case-insensitive Session storage semantics.
func TestSessionCaseInsensitiveContents(t *testing.T) {
	session := NewSessionWithID("abc")
	session.Set("UserName", NewApplicationString("lucas"))

	value, ok := session.Get("username")
	if !ok {
		t.Fatalf("expected key to exist")
	}
	if value.Type != ApplicationValueString || value.Str != "lucas" {
		t.Fatalf("unexpected value: %#v", value)
	}
}

// TestSessionProperties verifies SessionID numeric, timeout, LCID, and codepage behavior.
func TestSessionProperties(t *testing.T) {
	session := NewSessionWithID("sid-123")

	if session.SessionIDNumeric() <= 0 {
		t.Fatalf("expected positive numeric SessionID")
	}

	session.SetTimeout(0)
	if session.GetTimeout() != 20 {
		t.Fatalf("expected default timeout 20, got %d", session.GetTimeout())
	}

	session.SetLCID(0)
	if session.GetLCID() != resolveDefaultSessionLCID() {
		t.Fatalf("expected default LCID %d, got %d", resolveDefaultSessionLCID(), session.GetLCID())
	}

	session.SetCodePage(0)
	if session.GetCodePage() != 65001 {
		t.Fatalf("expected default codepage 65001, got %d", session.GetCodePage())
	}
}

// TestSessionPersistenceRoundtrip verifies disk persistence lifecycle for session data.
func TestSessionPersistenceRoundtrip(t *testing.T) {
	tempDir := t.TempDir()
	SetSessionStorageDir(filepath.Join(tempDir, "session"))
	t.Cleanup(func() { SetSessionStorageDir("") })

	session, err := CreateSession()
	if err != nil {
		t.Fatalf("CreateSession returned error: %v", err)
	}

	session.Set("Counter", NewApplicationInteger(9))
	session.SetTimeout(25)
	if err := session.Save(); err != nil {
		t.Fatalf("Save returned error: %v", err)
	}

	loaded, found, err := LoadSession(session.ID)
	if err != nil {
		t.Fatalf("LoadSession returned error: %v", err)
	}
	if !found {
		t.Fatalf("expected saved session to be found")
	}

	counter, ok := loaded.Get("counter")
	if !ok || counter.Type != ApplicationValueInteger || counter.Num != 9 {
		t.Fatalf("unexpected loaded counter value: %#v, found=%v", counter, ok)
	}
	if loaded.GetTimeout() != 25 {
		t.Fatalf("expected timeout 25, got %d", loaded.GetTimeout())
	}
}

// TestSessionStaticObjectsPersistenceRoundtrip verifies Session.StaticObjects survives disk roundtrip.
func TestSessionStaticObjectsPersistenceRoundtrip(t *testing.T) {
	tempDir := t.TempDir()
	SetSessionStorageDir(filepath.Join(tempDir, "session"))
	t.Cleanup(func() { SetSessionStorageDir("") })

	session, err := CreateSession()
	if err != nil {
		t.Fatalf("CreateSession returned error: %v", err)
	}

	session.AddStaticObject("SessObj", NewApplicationString("__AXON_STATIC_OBJECT_PROGID__:Scripting.Dictionary"))
	if err := session.Save(); err != nil {
		t.Fatalf("Save returned error: %v", err)
	}

	loaded, found, err := LoadSession(session.ID)
	if err != nil {
		t.Fatalf("LoadSession returned error: %v", err)
	}
	if !found {
		t.Fatalf("expected saved session to be found")
	}

	staticObj, ok := loaded.GetStaticObject("sessobj")
	if !ok {
		t.Fatalf("expected static object to be restored")
	}
	if staticObj.Type != ApplicationValueString || staticObj.Str != "__AXON_STATIC_OBJECT_PROGID__:Scripting.Dictionary" {
		t.Fatalf("unexpected restored static object value: %#v", staticObj)
	}
}

// TestSessionLockAndAbandon verifies lock tracking and abandon state behavior.
func TestSessionLockAndAbandon(t *testing.T) {
	session := NewSessionWithID("lock")

	session.Lock()
	session.Lock()
	if !session.IsLocked() || session.GetLockCount() != 2 {
		t.Fatalf("unexpected lock state: locked=%v count=%d", session.IsLocked(), session.GetLockCount())
	}

	session.Abandon()
	if !session.IsAbandoned() {
		t.Fatalf("expected session to be marked abandoned")
	}
	if session.Count() != 0 {
		t.Fatalf("expected empty contents after abandon")
	}

	session.Unlock()
	session.Unlock()
	if session.IsLocked() || session.GetLockCount() != 0 {
		t.Fatalf("expected unlocked state after unlock calls")
	}
}

// TestGetOrCreateSessionReusesRegisteredInstance verifies that repeated lookups for one
// ASP session ID reuse the same in-memory Session object.
func TestGetOrCreateSessionReusesRegisteredInstance(t *testing.T) {
	tempDir := t.TempDir()
	SetSessionStorageDir(filepath.Join(tempDir, "session"))
	t.Cleanup(func() { SetSessionStorageDir("") })

	session, err := CreateSession()
	if err != nil {
		t.Fatalf("CreateSession returned error: %v", err)
	}
	session.Set("Counter", NewApplicationInteger(7))
	if err := session.Save(); err != nil {
		t.Fatalf("Save returned error: %v", err)
	}

	first, isNew, err := GetOrCreateSession(session.ID)
	if err != nil {
		t.Fatalf("GetOrCreateSession returned error: %v", err)
	}
	if isNew {
		t.Fatalf("expected existing session to be reused")
	}

	second, isNew, err := GetOrCreateSession(session.ID)
	if err != nil {
		t.Fatalf("second GetOrCreateSession returned error: %v", err)
	}
	if isNew {
		t.Fatalf("expected second lookup to reuse existing session")
	}
	if first != second {
		t.Fatalf("expected the same in-memory Session pointer to be reused")
	}

	second.Set("Counter", NewApplicationInteger(11))
	value, ok := first.Get("counter")
	if !ok || value.Type != ApplicationValueInteger || value.Num != 11 {
		t.Fatalf("expected shared session mutation to be visible, got %#v found=%v", value, ok)
	}
}

// TestSessionSaveIfDirty verifies that SaveIfDirty skips disk writes when state is clean.
func TestSessionSaveIfDirty(t *testing.T) {
	tempDir := t.TempDir()
	SetSessionStorageDir(filepath.Join(tempDir, "session"))
	t.Cleanup(func() { SetSessionStorageDir("") })

	session := NewSessionWithID("dirty-check")
	if err := session.Save(); err != nil {
		t.Fatalf("Save returned error: %v", err)
	}
	if session.IsDirty() {
		t.Fatalf("expected clean session after Save")
	}

	path := sessionFilePath(session.ID)
	before, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("failed reading baseline session file: %v", err)
	}

	if err := session.SaveIfDirty(); err != nil {
		t.Fatalf("SaveIfDirty returned error: %v", err)
	}

	afterNoChange, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("failed reading unchanged session file: %v", err)
	}
	if string(before) != string(afterNoChange) {
		t.Fatalf("expected no file change when session is clean")
	}

	session.Set("Counter", NewApplicationInteger(42))
	if !session.IsDirty() {
		t.Fatalf("expected session to be dirty after mutation")
	}
	if err := session.SaveIfDirty(); err != nil {
		t.Fatalf("SaveIfDirty after mutation returned error: %v", err)
	}
	if session.IsDirty() {
		t.Fatalf("expected clean session after SaveIfDirty")
	}

	afterChange, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("failed reading changed session file: %v", err)
	}
	if string(before) == string(afterChange) {
		t.Fatalf("expected file change after dirty session flush")
	}
}

// TestFlushRegisteredSessions verifies force and dirty-only flush modes against registered sessions.
func TestFlushRegisteredSessions(t *testing.T) {
	tempDir := t.TempDir()
	SetSessionStorageDir(filepath.Join(tempDir, "session"))
	t.Cleanup(func() { SetSessionStorageDir("") })

	sessionRegistryMu.Lock()
	sessionRegistry = make(map[string]*Session)
	sessionRegistryMu.Unlock()

	session := NewSessionWithID("flush-1")
	registerSession(session)

	if err := FlushRegisteredSessions(false); err != nil {
		t.Fatalf("dirty flush failed: %v", err)
	}
	if session.IsDirty() {
		t.Fatalf("expected session clean after dirty flush")
	}

	session.Set("Name", NewApplicationString("axon"))
	if err := FlushRegisteredSessions(true); err != nil {
		t.Fatalf("forced flush failed: %v", err)
	}
	if session.IsDirty() {
		t.Fatalf("expected session clean after forced flush")
	}
}

// TestFlushRegisteredSessionsRemovesExpiredRegistered verifies expired registered sessions are deleted from disk and registry.
func TestFlushRegisteredSessionsRemovesExpiredRegistered(t *testing.T) {
	tempDir := t.TempDir()
	SetSessionStorageDir(filepath.Join(tempDir, "session"))
	t.Cleanup(func() { SetSessionStorageDir("") })

	sessionRegistryMu.Lock()
	sessionRegistry = make(map[string]*Session)
	sessionRegistryMu.Unlock()

	session := NewSessionWithID("expired-registered")
	session.SetTimeout(1)
	session.mu.Lock()
	session.LastAccessed = time.Now().Add(-2 * time.Minute)
	session.dirty = false
	session.mu.Unlock()

	if err := session.Save(); err != nil {
		t.Fatalf("Save returned error: %v", err)
	}
	registerSession(session)

	if err := FlushRegisteredSessions(false); err != nil {
		t.Fatalf("FlushRegisteredSessions returned error: %v", err)
	}

	if _, err := os.Stat(sessionFilePath(session.ID)); !os.IsNotExist(err) {
		t.Fatalf("expected expired registered session file to be deleted, got err=%v", err)
	}
	if existing := getRegisteredSession(session.ID); existing != nil {
		t.Fatalf("expected expired registered session to be unregistered")
	}
}

// TestFlushRegisteredSessionsRemovesExpiredOrphanedDiskSession verifies orphaned expired files are cleaned during flush cycle.
func TestFlushRegisteredSessionsRemovesExpiredOrphanedDiskSession(t *testing.T) {
	tempDir := t.TempDir()
	SetSessionStorageDir(filepath.Join(tempDir, "session"))
	t.Cleanup(func() { SetSessionStorageDir("") })

	sessionRegistryMu.Lock()
	sessionRegistry = make(map[string]*Session)
	sessionRegistryMu.Unlock()

	expired := NewSessionWithID("orphan-expired")
	expired.SetTimeout(1)
	expired.mu.Lock()
	expired.LastAccessed = time.Now().Add(-2 * time.Minute)
	expired.dirty = false
	expired.mu.Unlock()
	if err := expired.Save(); err != nil {
		t.Fatalf("Save expired returned error: %v", err)
	}

	active := NewSessionWithID("orphan-active")
	active.SetTimeout(20)
	active.mu.Lock()
	active.LastAccessed = time.Now()
	active.dirty = false
	active.mu.Unlock()
	if err := active.Save(); err != nil {
		t.Fatalf("Save active returned error: %v", err)
	}

	if err := FlushRegisteredSessions(false); err != nil {
		t.Fatalf("FlushRegisteredSessions returned error: %v", err)
	}

	if _, err := os.Stat(sessionFilePath(expired.ID)); !os.IsNotExist(err) {
		t.Fatalf("expected expired orphaned session file to be deleted, got err=%v", err)
	}
	if _, err := os.Stat(sessionFilePath(active.ID)); err != nil {
		t.Fatalf("expected active orphaned session file to remain, got err=%v", err)
	}
}
