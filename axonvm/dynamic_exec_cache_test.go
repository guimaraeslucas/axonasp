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
	"bytes"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

// TestExecuteCacheHitSameCompatibleScope verifies Execute dynamic compilation is reused in one compatible scope.
func TestExecuteCacheHitSameCompatibleScope(t *testing.T) {
	vm := buildEvalCacheTestVM(t)
	localSub := Value{Type: VTUserSub, Names: []string{"valueProbe"}}

	first, err := vm.getOrCompileDynamicProgram("valueProbe = 1", localSub, dynamicExecKindExecute)
	if err != nil {
		t.Fatalf("first compile failed: %v", err)
	}
	second, err := vm.getOrCompileDynamicProgram("valueProbe = 1", localSub, dynamicExecKindExecute)
	if err != nil {
		t.Fatalf("second compile failed: %v", err)
	}
	if first == nil || second == nil {
		t.Fatalf("expected non-nil execute cached payloads")
	}
	if first != second {
		t.Fatalf("expected execute cache hit to reuse same compiled payload")
	}
}

// TestExecuteCacheMissIncompatibleScope verifies Execute cache is invalidated when local scope shape changes.
func TestExecuteCacheMissIncompatibleScope(t *testing.T) {
	vm := buildEvalCacheTestVM(t)
	scopeA := Value{Type: VTUserSub, Names: []string{"valueProbe"}}
	scopeB := Value{Type: VTUserSub, Names: []string{"tempProbe", "valueProbe"}}

	first, err := vm.getOrCompileDynamicProgram("valueProbe = 1", scopeA, dynamicExecKindExecute)
	if err != nil {
		t.Fatalf("scopeA compile failed: %v", err)
	}
	second, err := vm.getOrCompileDynamicProgram("valueProbe = 1", scopeB, dynamicExecKindExecute)
	if err != nil {
		t.Fatalf("scopeB compile failed: %v", err)
	}
	if first == nil || second == nil {
		t.Fatalf("expected non-nil execute cached payloads")
	}
	if first == second {
		t.Fatalf("expected execute cache miss for incompatible local scope")
	}
}

// TestExecuteGlobalCacheHitStableClassMetadata verifies ExecuteGlobal reuses one cached fragment when scope metadata is stable.
func TestExecuteGlobalCacheHitStableClassMetadata(t *testing.T) {
	vm := buildEvalCacheTestVM(t)
	vm.registerRuntimeClass("Worker")

	first, err := vm.getOrCompileDynamicProgram("Dim g : g = 1", Value{}, dynamicExecKindExecuteGlobal)
	if err != nil {
		t.Fatalf("first compile failed: %v", err)
	}
	second, err := vm.getOrCompileDynamicProgram("Dim g : g = 1", Value{}, dynamicExecKindExecuteGlobal)
	if err != nil {
		t.Fatalf("second compile failed: %v", err)
	}
	if first == nil || second == nil {
		t.Fatalf("expected non-nil executeglobal cached payloads")
	}
	if first != second {
		t.Fatalf("expected executeglobal cache hit with stable class metadata")
	}
}

// TestExecuteCacheBypassInInteractiveMode verifies interactive CLI/TUI execution never reuses
// the process-wide Execute cache, ensuring edited input does not stick to stale compiled fragments.
func TestExecuteCacheBypassInInteractiveMode(t *testing.T) {
	vm := buildEvalCacheTestVM(t)
	vm.SetExecutionMode(ExecutionModeTUI)
	localSub := Value{Type: VTUserSub, Names: []string{"valueProbe"}}

	first, err := vm.getOrCompileDynamicProgram("valueProbe = 1", localSub, dynamicExecKindExecute)
	if err != nil {
		t.Fatalf("first interactive compile failed: %v", err)
	}
	second, err := vm.getOrCompileDynamicProgram("valueProbe = 1", localSub, dynamicExecKindExecute)
	if err != nil {
		t.Fatalf("second interactive compile failed: %v", err)
	}
	if first == nil || second == nil {
		t.Fatalf("expected non-nil interactive execute payloads")
	}
	if first == second {
		t.Fatalf("expected interactive execute path to bypass cache reuse")
	}

	globalFirst, err := vm.getOrCompileDynamicProgram("Dim g : g = 1", Value{}, dynamicExecKindExecuteGlobal)
	if err != nil {
		t.Fatalf("first interactive executeglobal compile failed: %v", err)
	}
	globalSecond, err := vm.getOrCompileDynamicProgram("Dim g : g = 1", Value{}, dynamicExecKindExecuteGlobal)
	if err != nil {
		t.Fatalf("second interactive executeglobal compile failed: %v", err)
	}
	if globalFirst == nil || globalSecond == nil {
		t.Fatalf("expected non-nil interactive executeglobal payloads")
	}
	if globalFirst == globalSecond {
		t.Fatalf("expected interactive executeglobal path to bypass cache reuse")
	}
}

// TestExecuteGlobalCachedFragmentReuseAvoidsBytecodeGrowth verifies repeated ExecuteGlobal calls reuse appended fragments.
func TestExecuteGlobalCachedFragmentReuseAvoidsBytecodeGrowth(t *testing.T) {
	vm := buildEvalCacheTestVM(t)
	host := NewMockHost()
	vm.SetHost(host)

	if _, err := vbsCompatExecuteGlobal(vm, []Value{NewString("Response.Write \"\"")}); err != nil {
		t.Fatalf("first ExecuteGlobal failed: %v", err)
	}
	lenAfterFirst := len(vm.bytecode)
	if _, err := vbsCompatExecuteGlobal(vm, []Value{NewString("Response.Write \"\"")}); err != nil {
		t.Fatalf("second ExecuteGlobal failed: %v", err)
	}
	lenAfterSecond := len(vm.bytecode)
	if lenAfterFirst != lenAfterSecond {
		t.Fatalf("expected cached ExecuteGlobal fragment to avoid bytecode growth: first=%d second=%d", lenAfterFirst, lenAfterSecond)
	}
}

// TestClassHeavyDynamicExecutionWithCacheReuse verifies class member Execute/Eval semantics remain correct with cache reuse.
func TestClassHeavyDynamicExecutionWithCacheReuse(t *testing.T) {
	source := `<%
Class Worker
    Public value
    Public Sub Init()
        value = 10
    End Sub
    Public Function RunOnce()
        Execute "value = value + 1"
        Execute "value = value + 1"
        RunOnce = Eval("value")
    End Function
End Class
Dim w
Set w = New Worker
w.Init
Response.Write CStr(w.RunOnce())
%>`

	compiler := NewASPCompiler(source)
	compiler.SetSourceName("/class_cache_test.asp")
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("run failed: %v", err)
	}
	host.Response().Flush()

	if strings.TrimSpace(output.String()) != "12" {
		t.Fatalf("unexpected class-heavy dynamic output: %q", output.String())
	}
	if len(vm.dynamicProgramStarts) == 0 {
		t.Fatalf("expected at least one cached dynamic fragment entry")
	}
}

// TestServerExecuteUsesCachedCompilationPath verifies nested Server.Execute reuses child-page cached compilation.
func TestServerExecuteUsesCachedCompilationPath(t *testing.T) {
	rootDir := t.TempDir()
	childPath := filepath.Join(rootDir, "child.asp")
	childCode := `<% Response.Write "C" %>`
	if err := os.WriteFile(childPath, []byte(childCode), 0o644); err != nil {
		t.Fatalf("write child failed: %v", err)
	}

	parentSource := `<%
Response.Write "P"
Server.Execute "child.asp"
Server.Execute "child.asp"
Response.Write "P"
%>`
	compiler := NewASPCompiler(parentSource)
	compiler.SetSourceName(filepath.Join(rootDir, "parent.asp"))
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/parent.asp")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "PCCP" {
		t.Fatalf("unexpected parent/child output: %q", output.String())
	}
	if host.ExecuteCompiles() != 1 {
		t.Fatalf("expected exactly one child compile, got %d", host.ExecuteCompiles())
	}
	if host.ExecuteCacheHits() < 1 {
		t.Fatalf("expected at least one child cache hit, got %d", host.ExecuteCacheHits())
	}
}

// BenchmarkDynamicExecuteHotPath measures steady-state Execute latency and allocations after warmup.
func BenchmarkDynamicExecuteHotPath(b *testing.B) {
	SetDumpPreprocessedSourceEnabled(false)
	_ = os.Unsetenv("DUMP_PREPROCESSED_SOURCE")

	compiler := NewASPCompiler(`<% Dim valueProbe : valueProbe = 0 %>`)
	compiler.SetSourceName("/bench_execute.asp")
	if err := compiler.Compile(); err != nil {
		b.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		b.Fatalf("vm run failed: %v", err)
	}

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		if _, err := vbsCompatExecute(vm, []Value{NewString("valueProbe = valueProbe + 1")}); err != nil {
			b.Fatalf("execute failed: %v", err)
		}
	}
}

// BenchmarkDynamicExecuteColdCompilePath approximates the pre-cache behavior by forcing unique source each iteration.
func BenchmarkDynamicExecuteColdCompilePath(b *testing.B) {
	SetDumpPreprocessedSourceEnabled(false)
	_ = os.Unsetenv("DUMP_PREPROCESSED_SOURCE")

	compiler := NewASPCompiler(`<% Dim valueProbe : valueProbe = 0 %>`)
	compiler.SetSourceName("/bench_execute_cold.asp")
	if err := compiler.Compile(); err != nil {
		b.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		b.Fatalf("vm run failed: %v", err)
	}

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		source := "valueProbe = valueProbe + 1 '" + strconv.Itoa(i)
		if _, err := vbsCompatExecute(vm, []Value{NewString(source)}); err != nil {
			b.Fatalf("execute failed: %v", err)
		}
	}
}
