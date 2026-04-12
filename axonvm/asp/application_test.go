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
	"testing"
	"time"
)

// TestApplicationCaseInsensitiveContents verifies case-insensitive storage and retrieval in Contents.
func TestApplicationCaseInsensitiveContents(t *testing.T) {
	app := NewApplication()
	app.Set("Counter", NewApplicationInteger(10))

	value, ok := app.Get("counter")
	if !ok {
		t.Fatalf("expected key to exist")
	}
	if value.Type != ApplicationValueInteger || value.Num != 10 {
		t.Fatalf("unexpected value: %#v", value)
	}
}

// TestApplicationLockNesting verifies nested lock and unlock bookkeeping.
func TestApplicationLockNesting(t *testing.T) {
	app := NewApplication()

	app.Lock()
	app.Lock()

	if !app.IsLocked() {
		t.Fatalf("expected application to be locked")
	}
	if app.GetLockCount() != 2 {
		t.Fatalf("expected lock count 2, got %d", app.GetLockCount())
	}

	app.Unlock()
	if !app.IsLocked() {
		t.Fatalf("expected application to remain locked after partial unlock")
	}
	if app.GetLockCount() != 1 {
		t.Fatalf("expected lock count 1, got %d", app.GetLockCount())
	}

	app.Unlock()
	if app.IsLocked() {
		t.Fatalf("expected application to be unlocked")
	}
	if app.GetLockCount() != 0 {
		t.Fatalf("expected lock count 0, got %d", app.GetLockCount())
	}
}

// TestApplicationCountAndRemove verifies count and removal behavior.
func TestApplicationCountAndRemove(t *testing.T) {
	app := NewApplication()
	app.Set("A", NewApplicationString("1"))
	app.Set("B", NewApplicationString("2"))
	app.AddStaticObject("C", NewApplicationBool(true))

	if app.Count() != 3 {
		t.Fatalf("expected count 3, got %d", app.Count())
	}

	app.Remove("a")
	if app.ContainsContent("A") {
		t.Fatalf("expected key A to be removed")
	}
	if app.Count() != 2 {
		t.Fatalf("expected count 2, got %d", app.Count())
	}

	app.RemoveAll()
	if app.Count() != 1 {
		t.Fatalf("expected only static object to remain, got count %d", app.Count())
	}
}

// TestApplicationLockBlocksOtherServers verifies Lock blocks other request servers until Unlock.
func TestApplicationLockBlocksOtherServers(t *testing.T) {
	app := NewApplication()
	owner := NewServer()
	other := NewServer()
	app.LockForServer(owner)

	released := make(chan struct{})
	go func() {
		app.WaitForServer(other)
		close(released)
	}()

	select {
	case <-released:
		t.Fatalf("expected second server to block while Application is locked")
	case <-time.After(100 * time.Millisecond):
	}

	app.UnlockForServer(owner)

	select {
	case <-released:
	case <-time.After(time.Second):
		t.Fatalf("expected second server to proceed after unlock")
	}
}
