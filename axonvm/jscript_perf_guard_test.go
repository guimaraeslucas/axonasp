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

import "testing"

var jscriptPerfGuardSink Value

// BenchmarkJScriptIntegerAddFastPath guards integer arithmetic throughput/allocation regressions.
func BenchmarkJScriptIntegerAddFastPath(b *testing.B) {
	vm := NewVM(nil, nil, 0)
	left := NewInteger(2147483000)
	right := NewInteger(7)

	b.ReportAllocs()
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		jscriptPerfGuardSink = vm.jsAdd(left, right)
	}
}

// BenchmarkJScriptIncrementFastPath guards ++ hot-path throughput/allocation regressions.
func BenchmarkJScriptIncrementFastPath(b *testing.B) {
	vm := NewVM(nil, nil, 0)
	value := NewInteger(123)

	b.ReportAllocs()
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		jscriptPerfGuardSink = vm.jsIncrementNumberValue(value)
	}
}

// TestJScriptIntegerFastPathBenchmarkGuard ensures integer hot paths stay allocation-free.
func TestJScriptIntegerFastPathBenchmarkGuard(t *testing.T) {
	addResult := testing.Benchmark(BenchmarkJScriptIntegerAddFastPath)
	incResult := testing.Benchmark(BenchmarkJScriptIncrementFastPath)

	if got := addResult.AllocsPerOp(); got != 0 {
		t.Fatalf("integer add fast path allocation regression: allocs/op=%d", got)
	}
	if got := incResult.AllocsPerOp(); got != 0 {
		t.Fatalf("increment fast path allocation regression: allocs/op=%d", got)
	}

	const maxAcceptableNsPerOp int64 = 20000
	if got := addResult.NsPerOp(); got > maxAcceptableNsPerOp {
		t.Fatalf("integer add fast path throughput regression: ns/op=%d limit=%d", got, maxAcceptableNsPerOp)
	}
	if got := incResult.NsPerOp(); got > maxAcceptableNsPerOp {
		t.Fatalf("increment fast path throughput regression: ns/op=%d limit=%d", got, maxAcceptableNsPerOp)
	}
}
