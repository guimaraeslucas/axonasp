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
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"testing"
)

// concurrentIncludeLibrary is the shared #include payload from the concurrency
// bug report: a class holding a Scripting.Dictionary plus functions that use
// RegExp and build object graphs.
const concurrentIncludeLibrary = `<%
Class Widget
	Private mItems
	Private Sub Class_Initialize()
		Set mItems = Server.CreateObject("Scripting.Dictionary")
	End Sub
	Public Sub AddItem(key, value)
		mItems(key) = value
	End Sub
	Public Function Describe()
		Dim k, s
		For Each k In mItems.Keys
			s = s & k & "=" & mItems(k) & ";"
		Next
		Describe = s
	End Function
End Class

Function Clean(text)
	Dim re
	Set re = New RegExp
	re.Pattern = "[^A-Za-z0-9 ]"
	re.Global = True
	Clean = re.Replace(text, "")
End Function

Function BuildWidget(n)
	Dim w, i
	Set w = New Widget
	For i = 1 To n
		w.AddItem "k" & i, Clean("value #" & i)
	Next
	Set BuildWidget = w
End Function
%>
`

// concurrentPageExpectedOutput is the value produced by Len(w.Describe()) for
// the 50 dictionary entries that the page builds.
const concurrentPageExpectedOutput = "632"

// concurrentPageNames lists the pages that share the include payload.
var concurrentPageNames = []string{"page1.asp", "page2.asp", "page3.asp"}

// concurrentPageSource builds one page source using the shared library.
func concurrentPageSource(pageName string) string {
	label := strings.TrimSuffix(pageName, filepath.Ext(pageName))
	return "<!--#include file=\"lib.inc\" -->\r\n<%\r\nDim w : Set w = BuildWidget(50)\r\nResponse.Write \"" +
		label + ": \" & Len(w.Describe())\r\n%>\r\n"
}

// TestScriptCacheConcurrentSharedIncludeExecution reproduces the concurrent
// request scenario from the bug report: several goroutines compile and execute
// pages that share one #include declaring a class, a Scripting.Dictionary and a
// RegExp. It must never trigger "concurrent map writes", must never leak
// compilation state between requests (which surfaced as error 800A01C2) and must
// always produce the same deterministic output.
//
// Run with -race to assert that cached program payloads stay read-only while VMs
// execute in parallel.
func TestScriptCacheConcurrentSharedIncludeExecution(t *testing.T) {
	modes := []struct {
		name string
		mode BytecodeCacheMode
	}{
		{name: "caching_enabled", mode: BytecodeCacheEnabled},
		{name: "caching_disabled", mode: BytecodeCacheDisabled},
	}

	for _, testCase := range modes {
		t.Run(testCase.name, func(t *testing.T) {
			t.Parallel()
			runConcurrentSharedIncludeScenario(t, testCase.mode)
		})
	}
}

// runConcurrentSharedIncludeScenario drives the concurrent execution scenario for
// one bytecode cache mode.
func runConcurrentSharedIncludeScenario(t *testing.T, mode BytecodeCacheMode) {
	t.Helper()

	const rounds = 40
	const executionsPerRound = 9

	rootDir := t.TempDir()
	writeConcurrentFixture(t, filepath.Join(rootDir, "lib.inc"), concurrentIncludeLibrary)
	for _, pageName := range concurrentPageNames {
		writeConcurrentFixture(t, filepath.Join(rootDir, pageName), concurrentPageSource(pageName))
	}

	cache := NewScriptCache(mode, filepath.Join(rootDir, "cache"), 16)

	type executionResult struct {
		pageName string
		output   string
		err      error
	}

	for round := range rounds {
		results := make([]executionResult, executionsPerRound)
		var workers sync.WaitGroup
		var ready sync.WaitGroup
		startGate := make(chan struct{})

		workers.Add(executionsPerRound)
		ready.Add(executionsPerRound)

		for slot := range executionsPerRound {
			go func(worker int) {
				defer workers.Done()

				pageName := concurrentPageNames[worker%len(concurrentPageNames)]
				pagePath := filepath.Join(rootDir, pageName)

				// Release every worker at the same instant so the cache and the
				// program payloads are hit by concurrent requests.
				ready.Done()
				<-startGate

				output, err := executeCachedASPPage(cache, pagePath, rootDir, pageName)
				results[worker] = executionResult{pageName: pageName, output: output, err: err}
			}(slot)
		}

		ready.Wait()
		close(startGate)
		workers.Wait()

		for slot, result := range results {
			if result.err != nil {
				t.Fatalf("round %d worker %d (%s) failed: %v", round, slot, result.pageName, result.err)
			}
			want := strings.TrimSuffix(result.pageName, filepath.Ext(result.pageName)) +
				": " + concurrentPageExpectedOutput
			if got := strings.TrimSpace(result.output); got != want {
				t.Fatalf("round %d worker %d (%s): got %q want %q", round, slot, result.pageName, got, want)
			}
		}
	}
}

// executeCachedASPPage compiles (or loads) the page through the shared cache and
// runs it on a pooled VM with an isolated ASP host.
func executeCachedASPPage(cache *ScriptCache, pagePath string, siteRoot string, virtualName string) (output string, err error) {
	defer func() {
		if recovered := recover(); recovered != nil {
			err = fmt.Errorf("panic during ASP execution: %v", recovered)
		}
	}()

	program, err := cache.LoadOrCompileWithOptions(pagePath, ScriptCompileOptions{IncludeSiteRoot: siteRoot})
	if err != nil {
		return "", err
	}

	vm := AcquireVMFromCachedProgram(program)
	defer vm.Release()

	host := NewMockHost()
	host.Server().SetRootDir(siteRoot)
	host.Server().SetRequestPath("/" + virtualName)

	var buffer bytes.Buffer
	host.SetOutput(&buffer)
	vm.SetHost(host)

	runErr := vm.Run()
	host.Response().Flush()
	if runErr != nil {
		return buffer.String(), runErr
	}
	return buffer.String(), nil
}

// writeConcurrentFixture writes one ASP fixture file, failing the test on error.
func writeConcurrentFixture(t *testing.T, path string, content string) {
	t.Helper()
	if err := os.WriteFile(path, []byte(content), 0o644); err != nil {
		t.Fatalf("write fixture %s failed: %v", path, err)
	}
}

// TestCachedProgramBytecodeStaysImmutableAfterExecution guards the invariant that
// cached program bytecode is read-only: resolving member calls must keep its
// inline-cache state inside the VM instead of patching the shared bytecode array,
// which is what leaked state between concurrent requests.
func TestCachedProgramBytecodeStaysImmutableAfterExecution(t *testing.T) {
	source := "<%\r\n" +
		"Class Counter\r\n" +
		"\tPrivate mValue\r\n" +
		"\tPublic Sub Add(delta)\r\n" +
		"\t\tmValue = mValue + delta\r\n" +
		"\tEnd Sub\r\n" +
		"\tPublic Function Total()\r\n" +
		"\t\tTotal = mValue\r\n" +
		"\tEnd Function\r\n" +
		"End Class\r\n" +
		"Dim c, i\r\n" +
		"Set c = New Counter\r\n" +
		"For i = 1 To 5\r\n" +
		"\tc.Add 3\r\n" +
		"Next\r\n" +
		"Response.Write \"total: \" & c.Total()\r\n" +
		"%>\r\n"

	rootDir := t.TempDir()
	pagePath := filepath.Join(rootDir, "counter.asp")
	writeConcurrentFixture(t, pagePath, source)

	cache := NewScriptCache(BytecodeCacheEnabled, filepath.Join(rootDir, "cache"), 16)
	program, err := cache.LoadOrCompileWithOptions(pagePath, ScriptCompileOptions{IncludeSiteRoot: rootDir})
	if err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	// The interpreter pool owns the program payload that every pooled VM in this
	// process shares, so it is the artifact that must stay read-only.
	pool := getProgramPool(program)
	poolBytecode := make([]byte, len(pool.program.Bytecode))
	copy(poolBytecode, pool.program.Bytecode)

	output, err := executeCachedASPPage(cache, pagePath, rootDir, "counter.asp")
	if err != nil {
		t.Fatalf("execution failed: %v", err)
	}
	if got := strings.TrimSpace(output); got != "total: 15" {
		t.Fatalf("unexpected output: got %q want %q", got, "total: 15")
	}

	if !bytes.Equal(pool.program.Bytecode, poolBytecode) {
		t.Fatalf("shared pooled program bytecode was mutated during execution")
	}

	cachedProgram, found := cache.getByCacheKey(normalizeScriptCacheKey(pagePath))
	if !found {
		t.Fatalf("cached program missing for %s", pagePath)
	}
	if !bytes.Equal(cachedProgram.Bytecode, poolBytecode) {
		t.Fatalf("cached program bytecode was mutated during execution")
	}
}
