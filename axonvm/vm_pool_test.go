//go:build !wasm

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
package axonvm

import (
	"bytes"
	"fmt"
	"sync"
	"testing"
	"time"
)

// TestAcquireVMFromCachedProgramResetsState verifies pooled VMs restore immutable program state
// and clear request-scoped data before being reused by another execution.
func TestAcquireVMFromCachedProgramResetsState(t *testing.T) {
	compiler := NewASPCompiler(`<% Response.Write "ok" %>`)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	program := cachedProgramFromCompiler(compiler)
	vm := AcquireVMFromCachedProgram(program)
	host := NewMockHost()
	vm.SetHost(host)
	vm.Globals[len(vm.Globals)-1] = NewString("dirty")
	vm.globalNames = append(vm.globalNames, "dynamicGlobal")
	vm.declaredGlobals["dynamicglobal"] = true
	vm.constGlobals["dynamicconst"] = true
	vm.responseCookieItems[20001] = "cookie"
	vm.nativeObjectProxies[20002] = nativeObjectProxy{ParentID: 1, Member: "Dirty"}
	vm.consoleTimerItems["dirty"] = time.Now()
	vm.ip = 9
	vm.sp = 3
	vm.lastLine = 42
	vm.Release()

	reused := AcquireVMFromCachedProgram(program)
	defer reused.Release()

	if reused.host != nil {
		t.Fatalf("expected pooled VM host to be cleared")
	}
	if reused.ip != 0 || reused.sp != -1 || reused.fp != 0 {
		t.Fatalf("expected VM execution pointers to be reset, got ip=%d sp=%d fp=%d", reused.ip, reused.sp, reused.fp)
	}
	if reused.lastLine != 0 || reused.lastColumn != 0 || reused.lastError != nil {
		t.Fatalf("expected last error state to be reset")
	}
	expectedGlobalCount := len(getBaseGlobalDictionary().names) + len(program.GlobalPreludeNames) + len(program.UserGlobalNames)
	if len(reused.globalNames) != expectedGlobalCount {
		t.Fatalf("expected global names to be restored, got %d want %d", len(reused.globalNames), expectedGlobalCount)
	}
	if _, exists := reused.declaredGlobals["dynamicglobal"]; exists {
		t.Fatalf("expected declared globals to be reset")
	}
	if _, exists := reused.constGlobals["dynamicconst"]; exists {
		t.Fatalf("expected const globals to be reset")
	}
	if len(reused.responseCookieItems) != 0 || len(reused.nativeObjectProxies) != 0 {
		t.Fatalf("expected dynamic native-object maps to be cleared")
	}
	if len(reused.consoleTimerItems) != 0 {
		t.Fatalf("expected console timer map to be cleared")
	}
	if reused.Globals[len(reused.Globals)-1].Type == VTString && reused.Globals[len(reused.Globals)-1].Str == "dirty" {
		t.Fatalf("expected globals to be restored from the base template")
	}
	if len(reused.Globals) >= 7 {
		if reused.Globals[0].Type != VTNativeObject || reused.Globals[0].Num != 0 {
			t.Fatalf("expected Response intrinsic to be restored")
		}
		if reused.Globals[4].Type != VTNativeObject || reused.Globals[4].Num != 4 {
			t.Fatalf("expected Application intrinsic to be restored")
		}
	}
}

// TestAcquireVMFromCachedProgramResetsJScriptState verifies pooled VM reuse clears
// JScript runtime state that can otherwise trigger stack underflow on expression pop.
func TestAcquireVMFromCachedProgramResetsJScriptState(t *testing.T) {
	compiler := NewASPCompiler(`<script runat="server" language="JScript">(1 + 2);</script>`)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	program := cachedProgramFromCompiler(compiler)
	vm := AcquireVMFromCachedProgram(program)

	vm.jsCallStack = append(vm.jsCallStack, jsCallFrame{returnIP: 1, savedSP: 3})
	vm.jsTryStack = append(vm.jsTryStack, 10)
	vm.jsErrStack = append(vm.jsErrStack, NewString("dirty"))
	vm.jsActiveEnvID = 99999
	vm.jsThisValue = NewString("dirty-this")
	vm.jsObjectItems[20001] = map[string]Value{"x": NewInteger(1)}
	vm.jsFunctionItems[20002] = &jsFunctionObject{name: "dirtyFn"}
	vm.jsForInItems[1] = &jsForInEnumerator{keys: []string{"k"}, index: 0}
	vm.jsForOfItems[2] = &jsForOfEnumerator{values: []Value{NewString("v")}, index: 0}
	vm.jsEnvItems[20003] = &jsEnvFrame{parentID: 0, bindings: map[string]Value{"x": NewInteger(1)}}

	vm.Release()

	reused := AcquireVMFromCachedProgram(program)
	defer reused.Release()

	if len(reused.jsCallStack) != 0 {
		t.Fatalf("expected jsCallStack to be reset")
	}
	if len(reused.jsTryStack) != 0 || len(reused.jsErrStack) != 0 {
		t.Fatalf("expected JScript try/error stacks to be reset")
	}
	if reused.jsActiveEnvID != 0 {
		t.Fatalf("expected jsActiveEnvID to be reset")
	}
	if reused.jsThisValue.Type != VTJSUndefined {
		t.Fatalf("expected jsThisValue to be undefined after reuse")
	}
	if len(reused.jsObjectItems) != 0 || len(reused.jsFunctionItems) != 0 || len(reused.jsForInItems) != 0 || len(reused.jsForOfItems) != 0 || len(reused.jsEnvItems) != 0 {
		t.Fatalf("expected JScript dynamic maps to be cleared")
	}

	host := NewMockHost()
	reused.SetHost(host)
	if err := reused.Run(); err != nil {
		t.Fatalf("expected reused VM run without stack underflow, got: %v", err)
	}
}

// TestAcquireVMFromCachedProgramPreservesOptionalDefaults ensures CachedProgram
// execution keeps Optional parameter default metadata (used by CLI/server paths).
func TestAcquireVMFromCachedProgramPreservesOptionalDefaults(t *testing.T) {
	compiler := NewASPCompiler(`<%
Function LogMessage(ByVal level As String, _
					ByVal msg As String, _
					Optional timestamp As String = "now", _
					ParamArray tags())
	Dim result, i
	result = "[" & level & "] " & msg
	If timestamp <> "now" Then
		result = result & " @" & timestamp
	End If
	If IsArray(tags) And UBound(tags) >= LBound(tags) Then
		result = result & " {"
		For i = LBound(tags) To UBound(tags)
			If i > LBound(tags) Then result = result & ", "
			result = result & tags(i)
		Next
		result = result & "}"
	End If
	LogMessage = result
End Function

Response.Write LogMessage("INFO", "Application started") & "|"
Response.Write LogMessage("WARN", "High memory usage", "now", "server1", "memory")
%>`)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	program := cachedProgramFromCompiler(compiler)
	vm := AcquireVMFromCachedProgram(program)
	defer vm.Release()

	var out bytes.Buffer
	host := NewMockHost()
	host.SetOutput(&out)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "[INFO] Application started|[WARN] High memory usage {server1, memory}"
	if got := out.String(); got != expected {
		t.Fatalf("cached optional defaults mismatch: expected %q, got %q", expected, got)
	}
}

func TestJScriptConcurrentPooledRunsNoStackUnderflow(t *testing.T) {
	compiler := NewASPCompiler(`<script runat="server" language="JScript">function id(v){return v;} var out=""; for (var i=0; i<10; i++) { out += id(i); } Response.Write(out);</script>`)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	program := cachedProgramFromCompiler(compiler)
	const workers = 24

	errCh := make(chan error, workers)
	var wg sync.WaitGroup
	wg.Add(workers)

	for i := 0; i < workers; i++ {
		go func() {
			defer wg.Done()
			vm := AcquireVMFromCachedProgram(program)
			defer vm.Release()

			host := NewMockHost()
			var output bytes.Buffer
			host.SetOutput(&output)
			host.Response().SetBuffer(false)
			vm.SetHost(host)

			if err := vm.Run(); err != nil {
				errCh <- err
				return
			}
			if output.String() != "0123456789" {
				errCh <- fmt.Errorf("unexpected JScript pooled output: %q", output.String())
			}
		}()
	}

	wg.Wait()
	close(errCh)

	for err := range errCh {
		if err != nil {
			t.Fatalf("concurrent pooled run failed: %v", err)
		}
	}
}

// TestCleanupRequestResourcesReleasesG3Image verifies pooled request cleanup drops image references.
func TestCleanupRequestResourcesReleasesG3Image(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	imageValue := vm.newG3ImageObject()
	imageObject := vm.g3imageItems[imageValue.Num]
	if imageObject == nil {
		t.Fatalf("expected g3image object to be allocated")
	}
	imageObject.DispatchMethod("new", []Value{NewInteger(32), NewInteger(32)})

	vm.CleanupRequestResources()

	if len(vm.g3imageItems) != 0 {
		t.Fatalf("expected g3image map to be cleared on request cleanup")
	}
	if imageObject.dc != nil || imageObject.lastLoaded != nil || imageObject.lastBytes != nil || imageObject.lastFontFace != nil {
		t.Fatalf("expected request cleanup to clear all g3image resource pointers")
	}
}

// TestAcquireVMFromCachedProgramClearsExtendedDynamicMaps verifies pooled VM reuse
// clears dynamic maps that can leak state across requests when not reset.
func TestAcquireVMFromCachedProgramClearsExtendedDynamicMaps(t *testing.T) {
	compiler := NewASPCompiler(`<% Response.Write "ok" %>`)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	program := cachedProgramFromCompiler(compiler)
	vm := AcquireVMFromCachedProgram(program)

	vm.g3searchItems[90000] = nil
	vm.g3stringBuilderItems[90001] = nil
	vm.g3zlibItems[90002] = nil
	vm.g3tarItems[90003] = nil
	vm.g3zstdItems[90004] = nil
	vm.g3axonliveItems[90005] = nil
	vm.g3axonliveProxyItems[90006] = nil
	vm.g3dbItems[90007] = nil
	vm.g3dbResultSetItems[90008] = nil
	vm.g3dbFieldsItems[90009] = nil
	vm.g3dbRowItems[90010] = nil
	vm.g3dbStatementItems[90011] = nil
	vm.g3dbTransactionItems[90012] = nil
	vm.g3dbResultItems[90013] = nil
	vm.runtimeClassItems[90014] = nil

	vm.jsRegExpItems[90001] = nil
	vm.jsArgumentsItems[90002] = nil
	vm.jsSetItems[90003] = nil
	vm.jsMapItems[90004] = nil
	vm.jsWeakRefItems[90005] = nil
	vm.jsFinalizationRegistryItems[90006] = nil
	vm.jsArrayIterators[90007] = nil
	vm.jsStringIterators[90008] = nil
	vm.jsRegExpStringIterators[90009] = nil
	vm.jsIntlCollatorItems[90010] = nil
	vm.jsIntlPluralRulesItems[90011] = nil
	vm.jsIntlRelativeTimeFormatItems[90012] = nil
	vm.jsProxyItems[90013] = nil

	if _, exists := vm.g3searchItems[90000]; !exists {
		t.Fatalf("expected g3search object to be allocated")
	}
	if _, exists := vm.g3stringBuilderItems[90001]; !exists {
		t.Fatalf("expected g3stringbuilder object to be allocated")
	}
	if _, exists := vm.g3zlibItems[90002]; !exists {
		t.Fatalf("expected g3zlib object to be allocated")
	}
	if _, exists := vm.g3tarItems[90003]; !exists {
		t.Fatalf("expected g3tar object to be allocated")
	}
	if _, exists := vm.g3zstdItems[90004]; !exists {
		t.Fatalf("expected g3zstd object to be allocated")
	}
	if _, exists := vm.g3axonliveItems[90005]; !exists {
		t.Fatalf("expected g3axonlive object to be allocated")
	}
	if _, exists := vm.g3dbItems[90007]; !exists {
		t.Fatalf("expected g3db object to be allocated")
	}
	if _, exists := vm.runtimeClassItems[90014]; !exists {
		t.Fatalf("expected runtime class object to be allocated")
	}

	vm.Release()

	reused := AcquireVMFromCachedProgram(program)
	defer reused.Release()

	if len(reused.g3searchItems) != 0 ||
		len(reused.g3stringBuilderItems) != 0 ||
		len(reused.g3zlibItems) != 0 ||
		len(reused.g3tarItems) != 0 ||
		len(reused.g3zstdItems) != 0 {
		t.Fatalf("expected core G3 maps to be cleared on pooled reuse")
	}
	if len(reused.g3axonliveItems) != 0 || len(reused.g3axonliveProxyItems) != 0 {
		t.Fatalf("expected G3AxonLive maps to be cleared on pooled reuse")
	}
	if len(reused.g3dbItems) != 0 ||
		len(reused.g3dbResultSetItems) != 0 ||
		len(reused.g3dbFieldsItems) != 0 ||
		len(reused.g3dbRowItems) != 0 ||
		len(reused.g3dbStatementItems) != 0 ||
		len(reused.g3dbTransactionItems) != 0 ||
		len(reused.g3dbResultItems) != 0 {
		t.Fatalf("expected G3DB maps to be cleared on pooled reuse")
	}
	if len(reused.runtimeClassItems) != 0 {
		t.Fatalf("expected runtime class map to be cleared on pooled reuse")
	}

	if len(reused.jsRegExpItems) != 0 ||
		len(reused.jsArgumentsItems) != 0 ||
		len(reused.jsSetItems) != 0 ||
		len(reused.jsMapItems) != 0 ||
		len(reused.jsWeakRefItems) != 0 ||
		len(reused.jsFinalizationRegistryItems) != 0 ||
		len(reused.jsArrayIterators) != 0 ||
		len(reused.jsStringIterators) != 0 ||
		len(reused.jsRegExpStringIterators) != 0 ||
		len(reused.jsIntlCollatorItems) != 0 ||
		len(reused.jsIntlPluralRulesItems) != 0 ||
		len(reused.jsIntlRelativeTimeFormatItems) != 0 ||
		len(reused.jsProxyItems) != 0 {
		t.Fatalf("expected extended JScript maps to be cleared on pooled reuse")
	}
}
