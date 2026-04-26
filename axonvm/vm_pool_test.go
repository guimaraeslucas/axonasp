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
	compiler := NewASPCompiler(`<script runat="server" language="JScript">foo();</script>`)
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
	if len(reused.jsObjectItems) != 0 || len(reused.jsFunctionItems) != 0 || len(reused.jsForInItems) != 0 || len(reused.jsEnvItems) != 0 {
		t.Fatalf("expected JScript dynamic maps to be cleared")
	}

	host := NewMockHost()
	reused.SetHost(host)
	if err := reused.Run(); err != nil {
		t.Fatalf("expected reused VM run without stack underflow, got: %v", err)
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
