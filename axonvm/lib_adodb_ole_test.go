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
	"runtime"
	"strings"
	"testing"
)

func resetADODBPlatformArchitectureConfigForTest(mode string) {
	adodbPlatformArchitectureTestOverride = mode
}

func TestADODBOLERecordsetOpenOptionsUseRecordsetSettings(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	rs := &adodbRecordset{
		cursorType: adOpenKeyset,
		lockType:   adLockOptimistic,
	}
	conn := &adodbConnection{cursorLocation: adUseServer}

	cursorLocation, cursorType, lockType := vm.adodbResolveOLERecordsetOpenOptions(rs, conn, nil)

	if cursorLocation != adUseServer {
		t.Fatalf("unexpected cursor location: %d", cursorLocation)
	}
	if cursorType != adOpenKeyset {
		t.Fatalf("unexpected cursor type: %d", cursorType)
	}
	if lockType != adLockOptimistic {
		t.Fatalf("unexpected lock type: %d", lockType)
	}
}

func TestADODBOLERecordsetOpenOptionsUpgradeClientCursor(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	rs := &adodbRecordset{
		cursorLocation: adUseClient,
		cursorType:     adOpenKeyset,
		lockType:       adLockOptimistic,
	}

	cursorLocation, cursorType, lockType := vm.adodbResolveOLERecordsetOpenOptions(rs, nil, nil)

	if cursorLocation != adUseClient {
		t.Fatalf("unexpected cursor location: %d", cursorLocation)
	}
	if cursorType != adOpenStatic {
		t.Fatalf("expected client cursor upgrade to static, got %d", cursorType)
	}
	if lockType != adLockOptimistic {
		t.Fatalf("unexpected lock type: %d", lockType)
	}
}

func TestADODBOLERecordsetOpenOptionsPreferExplicitOpenArgs(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	rs := &adodbRecordset{
		cursorLocation: adUseServer,
		cursorType:     adOpenForwardOnly,
		lockType:       adLockReadOnly,
	}
	args := []Value{
		NewString("SELECT 1"),
		NewEmpty(),
		NewInteger(adOpenDynamic),
		NewInteger(adLockBatchOptimistic),
	}

	_, cursorType, lockType := vm.adodbResolveOLERecordsetOpenOptions(rs, nil, args)

	if cursorType != adOpenDynamic {
		t.Fatalf("expected explicit cursor type override, got %d", cursorType)
	}
	if lockType != adLockBatchOptimistic {
		t.Fatalf("expected explicit lock type override, got %d", lockType)
	}
}

func TestADODBOLERecordsetOpenOptionsForceStaticForActiveConnection(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	rs := &adodbRecordset{
		activeConnection: 42,
		cursorType:       adOpenKeyset,
		lockType:         adLockOptimistic,
	}

	_, cursorType, lockType := vm.adodbResolveOLERecordsetOpenOptions(rs, nil, []Value{NewString("SELECT 1")})

	if cursorType != adOpenStatic {
		t.Fatalf("expected static cursor for ActiveConnection open, got %d", cursorType)
	}
	if lockType != adLockOptimistic {
		t.Fatalf("unexpected lock type: %d", lockType)
	}
}

func TestADODBOLEHydrateFieldMajorValues(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	rs := &adodbRecordset{columns: []string{"Id", "Name"}}
	values := []interface{}{int32(1), "alice", int32(2), "bob"}

	ok := vm.adodbHydrateRecordsetDataFromFieldMajorValues(rs, values, 2, 2)
	if !ok {
		t.Fatal("expected hydration success")
	}
	if len(rs.data) != 2 {
		t.Fatalf("unexpected row count: %d", len(rs.data))
	}
	if got := rs.data[0]["id"]; got.Type != VTInteger || got.Num != 1 {
		t.Fatalf("unexpected first row id: %#v", got)
	}
	if got := rs.data[0]["name"]; got.Type != VTString || got.Str != "alice" {
		t.Fatalf("unexpected first row name: %#v", got)
	}
	if got := rs.data[1]["id"]; got.Type != VTInteger || got.Num != 2 {
		t.Fatalf("unexpected second row id: %#v", got)
	}
	if got := rs.data[1]["name"]; got.Type != VTString || got.Str != "bob" {
		t.Fatalf("unexpected second row name: %#v", got)
	}
}

func TestADODBOLEHydrateFieldMajorValuesRejectsShortPayload(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	rs := &adodbRecordset{columns: []string{"Id", "Name"}}
	values := []interface{}{int32(1), "alice", int32(2)}

	ok := vm.adodbHydrateRecordsetDataFromFieldMajorValues(rs, values, 2, 2)
	if ok {
		t.Fatal("expected hydration failure for short payload")
	}
}

func TestADODBOLEPlatformProviderRewriteAutoToACE(t *testing.T) {
	defer resetADODBPlatformArchitectureConfigForTest("")
	resetADODBPlatformArchitectureConfigForTest("auto")
	input := "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\db\\site.mdb"
	got := adodbApplyPlatformAccessProvider(input)
	if runtime.GOARCH == "386" {
		if got != input {
			t.Fatalf("expected no rewrite on 386 auto, got %q", got)
		}
		return
	}
	if got == input {
		t.Fatalf("expected rewrite to ACE on non-386 auto")
	}
	if !strings.Contains(strings.ToLower(got), "provider=microsoft.ace.oledb.12.0") {
		t.Fatalf("expected ACE provider rewrite, got %q", got)
	}
}

func TestADODBOLEPlatformProviderRewriteForce386ToJET(t *testing.T) {
	defer resetADODBPlatformArchitectureConfigForTest("")
	resetADODBPlatformArchitectureConfigForTest("386")
	input := "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\db\\site.mdb"
	got := adodbApplyPlatformAccessProvider(input)
	if !strings.Contains(strings.ToLower(got), "provider=microsoft.jet.oledb.4.0") {
		t.Fatalf("expected Jet provider rewrite, got %q", got)
	}
}

func TestADODBOLEPlatformProviderRewriteForceAMD64ToACE(t *testing.T) {
	defer resetADODBPlatformArchitectureConfigForTest("")
	resetADODBPlatformArchitectureConfigForTest("amd64")
	input := "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\db\\site.mdb"
	got := adodbApplyPlatformAccessProvider(input)
	if !strings.Contains(strings.ToLower(got), "provider=microsoft.ace.oledb.12.0") {
		t.Fatalf("expected ACE provider rewrite, got %q", got)
	}
}

func TestADODBOLEPlatformProviderRewriteNonAccessUnchanged(t *testing.T) {
	defer resetADODBPlatformArchitectureConfigForTest("")
	resetADODBPlatformArchitectureConfigForTest("amd64")
	input := "Provider=SQLOLEDB;Server=localhost;Database=test"
	got := adodbApplyPlatformAccessProvider(input)
	if got != input {
		t.Fatalf("expected non-Access provider unchanged, got %q", got)
	}
}

func TestADODBOLEAlternateProviderSwapACEToJET(t *testing.T) {
	input := "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\db\\site.mdb"
	got, ok := adodbAlternateAccessProviderConnectionString(input)
	if !ok {
		t.Fatal("expected ACE provider fallback")
	}
	if !strings.Contains(strings.ToLower(got), "provider=microsoft.jet.oledb.4.0") {
		t.Fatalf("expected Jet fallback provider, got %q", got)
	}
}

func TestADODBOLEAlternateProviderSwapJETToACE(t *testing.T) {
	input := "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\db\\site.mdb"
	got, ok := adodbAlternateAccessProviderConnectionString(input)
	if !ok {
		t.Fatal("expected Jet provider fallback")
	}
	if !strings.Contains(strings.ToLower(got), "provider=microsoft.ace.oledb.12.0") {
		t.Fatalf("expected ACE fallback provider, got %q", got)
	}
}

func TestADODBOLEAlternateProviderNoSwapForNonAccess(t *testing.T) {
	input := "Provider=SQLOLEDB;Server=localhost;Database=test"
	got, ok := adodbAlternateAccessProviderConnectionString(input)
	if ok {
		t.Fatalf("expected no fallback provider, got %q", got)
	}
}
