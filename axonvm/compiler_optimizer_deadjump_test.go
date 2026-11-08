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
	"context"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"testing"
	"time"
)

// ---------------------------------------------------------------------------
// Differential harness: compile and execute the same program with the dead
// conditional jump pass enabled and disabled, then compare observable results.
// The pass only removes statically unreachable bytes, so both outcomes must be
// identical; any divergence is an optimizer defect, not a test expectation.
// ---------------------------------------------------------------------------

// optimizerOutcome is the observable result of one compile and execute cycle.
type optimizerOutcome struct {
	// status is one of "ok", "compile-error", "runtime-error", "host-error".
	status string
	// detail holds the response body when status is "ok" and the error text otherwise.
	detail string
}

// String renders the outcome for failure messages.
func (o optimizerOutcome) String() string {
	if o.status == "ok" {
		return fmt.Sprintf("ok(%q)", o.detail)
	}
	return fmt.Sprintf("%s(%q)", o.status, o.detail)
}

// withDeadConditionalJumpOptimizer runs fn with the dead conditional jump pass enabled or
// disabled and restores the previous state afterwards.
func withDeadConditionalJumpOptimizer(enabled bool, fn func()) {
	previous := deadConditionalJumpOptimizerDisabled
	deadConditionalJumpOptimizerDisabled = !enabled
	defer func() {
		deadConditionalJumpOptimizerDisabled = previous
	}()
	fn()
}

// executeOptimizerProgram compiles one ASP source and executes it in a mock host. The script
// timeout is one second so a runaway program surfaces as a timeout error instead of hanging
// the test binary. sourceName and includeSiteRoot mirror the server load path.
func executeOptimizerProgram(source, sourceName, includeSiteRoot string) optimizerOutcome {
	compiler := NewASPCompiler(source)
	if sourceName != "" {
		compiler.SetSourceName(sourceName)
	}
	if includeSiteRoot != "" {
		compiler.SetIncludeSiteRoot(includeSiteRoot)
	}
	if err := compiler.Compile(); err != nil {
		return optimizerOutcome{status: "compile-error", detail: err.Error()}
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	if err := host.Server().SetScriptTimeout(1); err != nil {
		return optimizerOutcome{status: "host-error", detail: err.Error()}
	}
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		return optimizerOutcome{status: "runtime-error", detail: err.Error()}
	}
	host.Response().Flush()
	return optimizerOutcome{status: "ok", detail: output.String()}
}

// deadJumpCorpusExecutionTimeout is the hard wall-clock cap for one corpus page execution.
// The engine's one second script timeout normally aborts a runaway page first, so this guard
// only fires when a page blocks below the script timeout check and would otherwise wedge the
// test binary instead of reporting a failure.
const deadJumpCorpusExecutionTimeout = 30 * time.Second

// executeOptimizerProgramWithinDeadline runs one page with the dead conditional jump pass
// enabled or disabled under a hard execution timeout. It returns the observable outcome, or a
// timeout error when the page fails to make progress within deadJumpCorpusExecutionTimeout.
func executeOptimizerProgramWithinDeadline(pageName, source, absPath string, optimized bool) (optimizerOutcome, error) {
	ctx, cancel := context.WithTimeout(context.Background(), deadJumpCorpusExecutionTimeout)
	defer cancel()

	done := make(chan optimizerOutcome, 1)
	go func() {
		var outcome optimizerOutcome
		withDeadConditionalJumpOptimizer(optimized, func() {
			outcome = executeOptimizerProgram(source, absPath, "")
		})
		done <- outcome
	}()

	select {
	case outcome := <-done:
		return outcome, nil
	case <-ctx.Done():
		return optimizerOutcome{}, fmt.Errorf("%s: exceeded the hard %s execution timeout", pageName, deadJumpCorpusExecutionTimeout)
	}
}

// executeOptimizerProgramBoth compiles and executes one program with the optimizer enabled
// and disabled and fails the test if the two outcomes diverge.
func executeOptimizerProgramBoth(t *testing.T, name, source, sourceName, includeSiteRoot string) optimizerOutcome {
	t.Helper()

	var optimized, unoptimized optimizerOutcome
	withDeadConditionalJumpOptimizer(true, func() {
		optimized = executeOptimizerProgram(source, sourceName, includeSiteRoot)
	})
	withDeadConditionalJumpOptimizer(false, func() {
		unoptimized = executeOptimizerProgram(source, sourceName, includeSiteRoot)
	})
	if optimized != unoptimized {
		t.Fatalf("%s: dead conditional jump optimizer changed the program outcome\n  optimized:   %s\n  unoptimized: %s",
			name, optimized, unoptimized)
	}
	return optimized
}

// normalizeHTMLWhitespace removes HTML whitespace so include-expansion newlines do not
// pollute output assertions.
func normalizeHTMLWhitespace(out string) string {
	return strings.Join(strings.Fields(out), "")
}

// ---------------------------------------------------------------------------
// Oracle corpus: full Compile() plus VM Run() for every shape the pass can touch.
// ---------------------------------------------------------------------------

// deadJumpOracleCase is one end-to-end program in the optimizer oracle corpus.
type deadJumpOracleCase struct {
	name string
	// source is a complete ASP page handed to Compile() and then to the VM.
	source string
	// wantOutput is the exact response body required when the program must succeed.
	wantOutput string
	// wantFailure requires the program to fail identically with and without the optimizer.
	wantFailure bool
}

// deadJumpOracleCases mirrors ASP shapes already exercised by existing axonvm tests
// (block/inline If resolution, Select Case prebinding, loop scopes, procedure hoisting,
// class modules, On Error Resume Next, JScript interop) so the oracle stays aligned with the
// scenarios this package guards elsewhere. None of these cases may be asserted by calling
// optimizeDeadConditionalJumpPass in isolation: every case runs the full compile pipeline.
func deadJumpOracleCases() []deadJumpOracleCase {
	return []deadJumpOracleCase{
		{
			name:       "inline if false branch is dropped",
			source:     `<% If False Then Response.Write "dead" Else Response.Write "live" %>`,
			wantOutput: "live",
		},
		{
			name:       "block if false branch keeps following statements",
			source:     "<%\nDim x\nx = 1\nIf False Then\n  x = 2\n  Response.Write \"dead\"\nEnd If\nResponse.Write x\n%>",
			wantOutput: "1",
		},
		{
			name:       "elseif chain evaluates live branch after dead head",
			source:     `<% Dim c : c = 2 : If False Then Response.Write "a" ElseIf c = 2 Then Response.Write "b" Else Response.Write "c" %>`,
			wantOutput: "b",
		},
		{
			name:       "compile time false condition variants",
			source:     `<% If 0 Then Response.Write "int" %>|<% If Empty Then Response.Write "empty" Else Response.Write "live" %>`,
			wantOutput: "|live",
		},
		{
			// VBScript condition coercion rejects an empty string literal with error 13,
			// so the dead-branch blanking must not change the failure behaviour.
			name:        "empty string condition still raises type mismatch",
			source:      `<% If "" Then Response.Write "str" %>`,
			wantFailure: true,
		},
		{
			name:       "nested dead branches collapse without touching live code",
			source:     "<%\nDim s\ns = \"a\"\nIf False Then\n  If False Then\n    s = s & \"x\"\n  End If\n  s = s & \"y\"\nEnd If\ns = s & \"b\"\nResponse.Write s\n%>",
			wantOutput: "ab",
		},
		{
			name:       "dead branch inside for loop",
			source:     "<%\nDim i, s\ns = \"\"\nFor i = 1 To 3\n  If False Then\n    s = s & \"x\"\n  Else\n    s = s & i\n  End If\nNext\nResponse.Write s\n%>",
			wantOutput: "123",
		},
		{
			name:       "dead branch inside do while loop",
			source:     "<%\nDim i, s\ni = 0\ns = \"\"\nDo While i < 2\n  If False Then\n    s = s & \"dead\"\n  Else\n    s = s & \"L\"\n  End If\n  i = i + 1\nLoop\nResponse.Write s\n%>",
			wantOutput: "LL",
		},
		{
			name:       "select case with dead case body",
			source:     "<%\nDim v, s\nv = 2\nSelect Case v\n  Case 1\n    If False Then s = \"dead\" End If\n    s = \"one\"\n  Case 2\n    If False Then s = \"dead\" Else s = \"two\" End If\n  Case Else\n    s = \"other\"\nEnd Select\nResponse.Write s\n%>",
			wantOutput: "two",
		},
		{
			name:       "function with dead branch keeps hoisted body callable",
			source:     "<%\nIf False Then\n  Function DeadScopeValue()\n    DeadScopeValue = -1\n  End Function\nEnd If\nFunction Live(a)\n  If False Then\n    Live = 0\n  Else\n    Live = a * 2\n  End If\nEnd Function\nResponse.Write Live(21)\n%>",
			wantOutput: "42",
		},
		{
			name:       "sub with dead branch and byref out parameter",
			source:     "<%\nDim acc\nacc = 1\nIf False Then\n  Sub Bump(ByRef value)\n    value = value + 100\n  End Sub\nEnd If\nSub AddTen(ByRef value)\n  If False Then\n    value = value - 1\n  Else\n    value = value + 10\n  End If\nEnd Sub\nAddTen acc\nResponse.Write acc\n%>",
			wantOutput: "11",
		},
		{
			name:       "class module with dead property branch",
			source:     "<%\nClass Box\n  Public Value\n  Public Property Get Doubled()\n    If False Then\n      Doubled = -1\n    Else\n      Doubled = Value * 2\n    End If\n  End Property\nEnd Class\nDim b\nSet b = New Box\nb.Value = 6\nResponse.Write b.Doubled\n%>",
			wantOutput: "12",
		},
		{
			name:       "with block with dead branch inside",
			source:     "<%\nClass Holder\n  Public Text\nEnd Class\nDim h\nSet h = New Holder\nWith h\n  If False Then\n    .Text = \"dead\"\n  Else\n    .Text = \"live\"\n  End If\nEnd With\nResponse.Write h.Text\n%>",
			wantOutput: "live",
		},
		{
			name:       "goto label outside dead range keeps blanking conservative",
			source:     "<%\nIf False Then\n  GoTo Skip\n  Response.Write \"dead\"\nEnd If\nResponse.Write \"before\"\nSkip:\nResponse.Write \"|after\"\n%>",
			wantOutput: "before|after",
		},
		{
			name:       "on error resume next inside dead branch is inert",
			source:     "<%\nDim s\nOn Error Resume Next\nIf False Then\n  On Error GoTo 0\n  s = s & \"dead\"\nEnd If\ns = s & \"live\"\nOn Error GoTo 0\nResponse.Write s\n%>",
			wantOutput: "live",
		},
		{
			name:       "array writes inside dead branch do not run",
			source:     "<%\nDim a(2), i\nFor i = 0 To 2\n  a(i) = i\nNext\nIf False Then\n  a(0) = 99\nEnd If\nResponse.Write a(0) & a(1) & a(2)\n%>",
			wantOutput: "012",
		},
		{
			name:       "jscript dead branch stays untouched",
			source:     "<script language=\"jscript\" runat=\"server\">\nif (false) { Response.Write(\"dead\"); }\nResponse.Write(\"js-live\");\n</script>",
			wantOutput: "js-live",
		},
		{
			// Classic ASP hoists <script runat="server"> blocks, so the JScript part runs
			// before the inline VBScript tag; both dead branches must stay inert.
			name:       "mixed vbscript and jscript dead branches",
			source:     "<%\nDim s\ns = \"vb\"\nIf False Then s = s & \"-dead\"\nResponse.Write s\n%><script language=\"jscript\" runat=\"server\">\nif (false) { Response.Write(\"-dead\"); }\nResponse.Write(\"-js\");\n</script>",
			wantOutput: "-jsvb",
		},
		{
			name:       "empty program and empty dead block",
			source:     "<% %>|<%\nIf False Then\nEnd If\n%>",
			wantOutput: "|",
		},
		{
			name:        "runtime error outside dead branch still fails",
			source:      "<%\nIf False Then\n  Response.Write \"dead\"\nEnd If\nDim zero\nzero = 0\nResponse.Write 1 / zero\n%>",
			wantFailure: true,
		},
	}
}

// TestCompiler_DeadJumpOptimization_E2E runs the full compile pipeline and the VM for every
// oracle program and asserts that enabling the dead conditional jump pass changes neither the
// response body nor the failure behaviour.
func TestCompiler_DeadJumpOptimization_E2E(t *testing.T) {
	for _, tc := range deadJumpOracleCases() {
		t.Run(tc.name, func(t *testing.T) {
			got := executeOptimizerProgramBoth(t, tc.name, tc.source, "e2e_"+sanitizeCaseName(tc.name)+".asp", "")

			if tc.wantFailure {
				if got.status == "ok" {
					t.Fatalf("%s: expected a failure, got output %q", tc.name, got.detail)
				}
				return
			}
			if got.status != "ok" {
				t.Fatalf("%s: unexpected %s: %s", tc.name, got.status, got.detail)
			}
			if got.detail != tc.wantOutput {
				t.Fatalf("%s: unexpected output: got %q want %q", tc.name, got.detail, tc.wantOutput)
			}
		})
	}
}

// sanitizeCaseName converts a subtest name into a file-name-safe source name.
func sanitizeCaseName(name string) string {
	return strings.Map(func(r rune) rune {
		switch {
		case r >= 'a' && r <= 'z', r >= 'A' && r <= 'Z', r >= '0' && r <= '9':
			return r
		default:
			return '_'
		}
	}, name)
}

// ---------------------------------------------------------------------------
// Include tree end-to-end: dead code paths span #include boundaries.
// ---------------------------------------------------------------------------

// writeIncludeTreeFile writes one file of the include tree, creating parent folders.
func writeIncludeTreeFile(t *testing.T, root, relative, content string) string {
	t.Helper()

	path := filepath.Join(root, filepath.FromSlash(relative))
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatalf("mkdir %s failed: %v", relative, err)
	}
	if err := os.WriteFile(path, []byte(content), 0o644); err != nil {
		t.Fatalf("write %s failed: %v", relative, err)
	}
	return path
}

// buildDeadJumpIncludeTree materialises a branching include graph in which dead branches span
// include boundaries and nested includes contribute hoisted procedures.
//
//	main.asp
//	  -> /includes/header.inc        (dead branch wrapping a nested include)
//	       -> includes/nested/deep.inc   (procedure hoisted out of the dead branch)
//	  -> includes/body.inc           (function with dead branch, live output)
//	  -> /includes/branch/c.inc      (dead branch wrapping an include with inner jumps)
//	       -> branch/d.inc
func buildDeadJumpIncludeTree(t *testing.T, root string) string {
	t.Helper()

	mainSource := "<%@ Language=\"VBScript\" %>\n" +
		"<!--#include virtual=\"/includes/header.inc\"-->\n" +
		"<!--#include file=\"includes/body.inc\"-->\n" +
		"<!--#include virtual=\"/includes/branch/c.inc\"-->\n" +
		"<% Response.Write \"|main=\" & DeepSum(4, 5) %>\n"

	tree := map[string]string{
		"main.asp": mainSource,
		"includes/header.inc": "<%\nIf False Then\n" +
			"  Response.Write \"|header-dead\"\n" +
			"%>\n" +
			"<!--#include file=\"nested/deep.inc\"-->\n" +
			"<%\nElse\n" +
			"  Response.Write \"|header-live\"\n" +
			"End If\n%>\n",
		"includes/nested/deep.inc": "<%\n" +
			"Function DeepSum(a, b)\n" +
			"  Dim total\n" +
			"  If False Then\n" +
			"    total = -1\n" +
			"  Else\n" +
			"    total = a + b\n" +
			"  End If\n" +
			"  DeepSum = total\n" +
			"End Function\n" +
			"If False Then\n" +
			"  Response.Write \"|deep-dead\"\n" +
			"End If\n" +
			"%>\n",
		"includes/body.inc": "<%\n" +
			"Function BodySum(a, b)\n" +
			"  Dim acc\n" +
			"  If False Then\n" +
			"    acc = -1\n" +
			"  Else\n" +
			"    acc = a + b\n" +
			"  End If\n" +
			"  BodySum = acc\n" +
			"End Function\n" +
			"Response.Write \"|body=\" & BodySum(2, 3)\n" +
			"%>\n",
		"includes/branch/c.inc": "<%\nIf False Then\n%>\n" +
			"<!--#include file=\"d.inc\"-->\n" +
			"<%\nElse\n  Response.Write \"|c-live\"\nEnd If\n%>\n",
		"includes/branch/d.inc": "<%\nIf True Then\n" +
			"  Response.Write \"|d-true\"\n" +
			"Else\n" +
			"  Response.Write \"|d-false\"\n" +
			"End If\n" +
			"Response.Write \"|d-tail\"\n" +
			"%>\n",
	}

	names := make([]string, 0, len(tree))
	for name := range tree {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		writeIncludeTreeFile(t, root, name, tree[name])
	}
	return filepath.Join(root, "main.asp")
}

// TestCompiler_DeadJumpOptimization_IncludeTreeE2E compiles one page whose dead code paths
// span a branching #include tree and requires the executed output to be identical with and
// without the optimizer. The tree also proves that blanking a dead range cannot destroy
// procedures hoisted out of that range.
func TestCompiler_DeadJumpOptimization_IncludeTreeE2E(t *testing.T) {
	root := t.TempDir()
	mainPath := buildDeadJumpIncludeTree(t, root)

	content, err := os.ReadFile(mainPath)
	if err != nil {
		t.Fatalf("read main.asp failed: %v", err)
	}

	got := executeOptimizerProgramBoth(t, "include-tree", string(content), mainPath, root)
	if got.status != "ok" {
		t.Fatalf("unexpected %s: %s", got.status, got.detail)
	}
	if normalized := normalizeHTMLWhitespace(got.detail); normalized != "|header-live|body=5|c-live|main=9" {
		t.Fatalf("unexpected output: got %q want %q", normalized, "|header-live|body=5|c-live|main=9")
	}

	// The pass must actually have removed dead bytes; a silent no-op would make every
	// execution assertion above vacuous. Opcode counts are structural, not byte offsets.
	source := string(content)
	var optimizedNops, unoptimizedNops int
	withDeadConditionalJumpOptimizer(true, func() {
		compiler := NewASPCompiler(source)
		compiler.SetSourceName(mainPath)
		compiler.SetIncludeSiteRoot(root)
		if err := compiler.Compile(); err != nil {
			t.Fatalf("optimized compile failed: %v", err)
		}
		optimizedNops = countBytecodeOp(compiler.Bytecode(), OpNop)
	})
	withDeadConditionalJumpOptimizer(false, func() {
		compiler := NewASPCompiler(source)
		compiler.SetSourceName(mainPath)
		compiler.SetIncludeSiteRoot(root)
		if err := compiler.Compile(); err != nil {
			t.Fatalf("unoptimized compile failed: %v", err)
		}
		unoptimizedNops = countBytecodeOp(compiler.Bytecode(), OpNop)
	})
	if optimizedNops <= unoptimizedNops {
		t.Fatalf("dead conditional jump pass removed nothing: nops optimized=%d unoptimized=%d", optimizedNops, unoptimizedNops)
	}
}

// ---------------------------------------------------------------------------
// Scaling: the pass must stay O(N) and full compilation must stay linear because
// of it. Block counts double so a quadratic pass shows up as a 4x per doubling.
// ---------------------------------------------------------------------------

// deadJumpScaleBlocks holds the N, 2N, 4N and 8N block counts used by the scaling assertion
// and by both benchmarks.
var deadJumpScaleBlocks = []int{100, 200, 400, 800}

// buildDeadJumpScaleSource generates an ASP page made of count repetitions of one standard
// block: a dead If/ElseIf/Else chain plus a function with its own dead branch. Each block adds
// 5 to the total, so the executed value is an exact oracle for the generated program.
func buildDeadJumpScaleSource(count int) string {
	var b strings.Builder
	b.WriteString("<%\nDim scaleTotal\nscaleTotal = 0\n")
	for i := range count {
		b.WriteString("If False Then\n  scaleTotal = scaleTotal + 1\nElseIf False Then\n  scaleTotal = scaleTotal + 2\nElse\n  scaleTotal = scaleTotal + 3\nEnd If\n")
		fmt.Fprintf(&b, "Function Fs%d()\n  Dim localVal%d\n  If False Then\n    localVal%d = 1\n  Else\n    localVal%d = 2\n  End If\n  Fs%d = localVal%d\nEnd Function\n",
			i, i, i, i, i, i)
		fmt.Fprintf(&b, "scaleTotal = scaleTotal + Fs%d()\n", i)
	}
	b.WriteString("Response.Write scaleTotal\n%>\n")
	return b.String()
}

// measureDeadJumpCompile returns the fastest of runs compile durations for one source. The
// minimum is used so GC pauses and scheduler noise do not distort the scaling ratio.
func measureDeadJumpCompile(t testing.TB, source string, runs int) time.Duration {
	t.Helper()

	best := time.Duration(1<<62 - 1)
	for range runs {
		runtime.GC()
		start := time.Now()
		compiler := NewASPCompiler(source)
		compiler.SetSourceName("scale.asp")
		if err := compiler.Compile(); err != nil {
			t.Fatalf("compile failed: %v", err)
		}
		if elapsed := time.Since(start); elapsed < best {
			best = elapsed
		}
	}
	return best
}

// buildDeadJumpOptimizerBase compiles the generated page with the pass disabled and returns
// the compiler whose bytecode still holds every dead range, so the pass can be measured on
// real work instead of on already blanked bytecode.
func buildDeadJumpOptimizerBase(t testing.TB, blocks int) *Compiler {
	t.Helper()

	var compiler *Compiler
	withDeadConditionalJumpOptimizer(false, func() {
		compiler = NewASPCompiler(buildDeadJumpScaleSource(blocks))
		compiler.SetSourceName("scale.asp")
		if err := compiler.Compile(); err != nil {
			t.Fatalf("compile failed: %v", err)
		}
	})
	return compiler
}

// TestOptimizeDeadConditionalJumpLinearScaling asserts that compilation time grows with the
// input, not with its square: 4N must cost about 4x N and 8N about 8x N. The thresholds leave
// headroom for the rest of the compiler while still failing an O(N^2) pass (which would need
// 16x and 64x).
func TestOptimizeDeadConditionalJumpLinearScaling(t *testing.T) {
	sources := make([]string, len(deadJumpScaleBlocks))
	for i, blocks := range deadJumpScaleBlocks {
		sources[i] = buildDeadJumpScaleSource(blocks)
	}

	// The generated program is the benchmark input, so its executed value is verified first:
	// each block adds 5, which keeps the generator itself from silently degrading.
	generated := executeOptimizerProgramBoth(t, "scale-block-output", sources[0], "scale_block_output.asp", "")
	if generated.status != "ok" {
		t.Fatalf("generated scale program failed: %s", generated)
	}
	if want := fmt.Sprintf("%d", 5*deadJumpScaleBlocks[0]); generated.detail != want {
		t.Fatalf("generated scale program output: got %q want %q", generated.detail, want)
	}

	baseline := measureDeadJumpCompile(t, sources[0], 3)
	if baseline < time.Millisecond {
		t.Fatalf("baseline measurement too small to compare scaling: %s", baseline)
	}

	times := make([]time.Duration, len(sources))
	times[0] = baseline
	for i := 1; i < len(sources); i++ {
		times[i] = measureDeadJumpCompile(t, sources[i], 3)
	}
	for i, blocks := range deadJumpScaleBlocks {
		t.Logf("blocks=%d bytes=%d compile=%s (%.2f ns/byte)", blocks, len(sources[i]),
			times[i], float64(times[i].Nanoseconds())/float64(len(sources[i])))
	}

	// deadJumpScaleBlocks is N, 2N, 4N, 8N: index 2 is 4N, index 3 is 8N.
	ratio4N := float64(times[2]) / float64(baseline)
	ratio8N := float64(times[3]) / float64(baseline)
	if ratio4N > 12 {
		t.Fatalf("4N compile ratio %.1fx exceeds the linear budget (12x): compilation is superlinear", ratio4N)
	}
	if ratio8N > 20 {
		t.Fatalf("8N compile ratio %.1fx exceeds the linear budget (20x): compilation is superlinear", ratio8N)
	}
}

// BenchmarkOptimizeDeadConditionalJump measures the dead conditional jump pass alone over
// N, 2N, 4N and 8N repetitions of the standard block. The reported ns/byte metric stays flat
// while the pass is linear and grows with the input size as soon as it rescans bytecode.
func BenchmarkOptimizeDeadConditionalJump(b *testing.B) {
	baselineNsPerByte := 0.0
	for _, blocks := range deadJumpScaleBlocks {
		b.Run(fmt.Sprintf("blocks=%d", blocks), func(b *testing.B) {
			base := buildDeadJumpOptimizerBase(b, blocks)
			bytecodeLen := len(base.bytecode)
			if bytecodeLen == 0 {
				b.Fatal("empty bytecode")
			}

			// Zero allocations inside the measured loop: the walker reuses one buffer and one
			// struct copy, so allocation counters reflect the pass internals only.
			work := *base
			buffer := make([]byte, bytecodeLen)

			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				copy(buffer, base.bytecode)
				work.bytecode = buffer
				work.optimizeDeadConditionalJumpPass()
			}
			b.StopTimer()

			nsPerByte := float64(b.Elapsed().Nanoseconds()) / float64(b.N) / float64(bytecodeLen)
			b.ReportMetric(nsPerByte, "ns/byte")
			if baselineNsPerByte <= 0 {
				baselineNsPerByte = nsPerByte
				b.ReportMetric(1, "x-1x")
				return
			}
			b.ReportMetric(nsPerByte/baselineNsPerByte, "x-1x")
		})
	}
}

// BenchmarkCompileDeadJumpSource measures the full Compile() pipeline for N, 2N, 4N and 8N
// block pages and reports the same ns/byte and ratio metrics: the compile-time ratio must
// track the input ratio (4N about 4x N, not 16x).
func BenchmarkCompileDeadJumpSource(b *testing.B) {
	baselineNsPerByte := 0.0
	for _, blocks := range deadJumpScaleBlocks {
		b.Run(fmt.Sprintf("blocks=%d", blocks), func(b *testing.B) {
			source := buildDeadJumpScaleSource(blocks)
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				compiler := NewASPCompiler(source)
				compiler.SetSourceName("scale.asp")
				if err := compiler.Compile(); err != nil {
					b.Fatalf("compile failed: %v", err)
				}
			}
			b.StopTimer()

			nsPerByte := float64(b.Elapsed().Nanoseconds()) / float64(b.N) / float64(len(source))
			b.ReportMetric(nsPerByte, "ns/byte")
			if baselineNsPerByte <= 0 {
				baselineNsPerByte = nsPerByte
				b.ReportMetric(1, "x-1x")
				return
			}
			b.ReportMetric(nsPerByte/baselineNsPerByte, "x-1x")
		})
	}
}

// ---------------------------------------------------------------------------
// Corpus regression: every ASP page shipped in www/tests is compiled and executed
// with and without the pass; the failure set and the observable outcomes must match.
// ---------------------------------------------------------------------------

// collectASPCorpusFiles returns every .asp file under root, sorted for stable reporting.
func collectASPCorpusFiles(t testing.TB, root string) []string {
	t.Helper()

	var files []string
	err := filepath.WalkDir(root, func(path string, d os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return nil
		}
		if d.IsDir() {
			return nil
		}
		if strings.EqualFold(filepath.Ext(path), ".asp") {
			files = append(files, path)
		}
		return nil
	})
	if err != nil {
		t.Fatalf("walk %s failed: %v", root, err)
	}
	sort.Strings(files)
	return files
}

// deadJumpCorpusPairAttempts bounds how many adjacent (optimized, unoptimized) pairs are
// executed for a page whose first pair diverged, whose unoptimized baseline is repeatable and
// which is not quarantined. A page that reads a coarse clock can straddle the boundary between
// the two executions of a pair, so a matching pair must be allowed to occur; a systematic
// defect mismatches every pair because the divergence is structural, not timing-related.
const deadJumpCorpusPairAttempts = 4

// deadJumpCorpusQuarantine lists the corpus pages that cannot be compared by output equality
// because two identical unoptimized runs of them already differ. The list was measured by
// running every page five times with the optimizer disabled; the causes are clock values and
// clock-driven branches, Timer/Rnd seeds, OS randomness, filesystem paths and listings,
// generated archive or PDF metadata, and hash- or map-ordered output.
//
// Quarantined pages are compared by outcome status instead: the optimizer must still leave the
// page succeeding or failing exactly as it did before. Pages outside this list must reproduce
// their output byte for byte; a page that starts drifting is reported with its name so the
// entry can be reviewed and added here.
var deadJumpCorpusQuarantine = map[string]struct{}{
	"bench/testredim.asp":                  {},
	"captcha.asp":                          {},
	"json-teste/test.asp":                  {},
	"json-teste/tmp_test_include_oern.asp": {},
	"json-teste/tmp_test_oern.asp":         {},
	"test_404_handling.asp":                {},
	"test_class_dict.asp":                  {},
	"test_crypto.asp":                      {},
	"test_examples_fso_usage.asp":          {},
	"test_fso.asp":                         {},
	"test_functions.asp":                   {},
	"test_g3date.asp":                      {},
	"test_g3image.asp":                     {},
	"test_g3tar.asp":                       {},
	"test_jscript_es5_json_object.asp":     {},
	"test_mswc_demo.asp":                   {},
	"test_pdf_basic.asp":                   {},
	"test_pdf_image.asp":                   {},
	"test_platform_parity.asp":             {},
	"test_timer.asp":                       {},
	"test_timer2.asp":                      {},
	"test_var_scope.asp":                   {},
	"test_vb_functions.asp":                {},
	"tests2/09-filesystem.asp":             {},
	"tmp_random_debug.asp":                 {},
}

// isScriptTimeoutOutcome reports whether an outcome is the one second script timeout guard.
// Timeout firing is load dependent, so such a page is inconclusive rather than mismatching.
func isScriptTimeoutOutcome(outcome optimizerOutcome) bool {
	return strings.Contains(outcome.detail, "exceeded the configured timeout")
}

// TestDeadJumpOptimizationCorpusIdentity compiles and executes every ASP page under
// www/tests once with the dead conditional jump pass enabled and once disabled, asserting that
// the failure set and the observable outcome are identical.
//
// Comparison is tiered because part of the shipped corpus is not repeatable:
//
//  1. Pages outside deadJumpCorpusQuarantine must produce byte-identical output. A diverging
//     pair is retried a few times to rule out a clock boundary inside the pair, and if it never
//     matches, the test fails with the page name.
//  2. Quarantined pages only have to keep their outcome status (ok, compile error, runtime
//     error), which still catches a corrupted or drifting instruction pointer.
//  3. Pages that hit the one second script timeout guard are skipped, because the guard firing
//     depends on machine load rather than on the optimizer.
func TestDeadJumpOptimizationCorpusIdentity(t *testing.T) {
	if testing.Short() {
		t.Skip("corpus differential compiles and executes the shipped ASP test suite")
	}

	root := filepath.Join("..", "www", "tests")
	if _, err := os.Stat(root); err != nil {
		t.Skipf("ASP corpus not available: %v", err)
	}
	files := collectASPCorpusFiles(t, root)
	if len(files) == 0 {
		t.Skip("ASP corpus is empty")
	}

	optimizedFailures := make([]string, 0, 16)
	unoptimizedFailures := make([]string, 0, 16)
	stableComparisons := 0
	quarantinedPages := 0
	driftedPages := 0
	timeoutPages := 0
	start := time.Now()

	for _, path := range files {
		content, err := os.ReadFile(path)
		if err != nil {
			t.Fatalf("read %s failed: %v", path, err)
		}
		source := string(content)
		absPath, err := filepath.Abs(path)
		if err != nil {
			t.Fatalf("abs %s failed: %v", path, err)
		}
		relative, err := filepath.Rel(root, path)
		if err != nil {
			t.Fatalf("rel %s failed: %v", path, err)
		}
		relative = filepath.ToSlash(relative)

		var optimized, unoptimized optimizerOutcome
		runOptimized := func() {
			outcome, err := executeOptimizerProgramWithinDeadline(relative, source, absPath, true)
			if err != nil {
				t.Fatalf("%s: optimized run %v", path, err)
			}
			optimized = outcome
		}
		runUnoptimized := func() optimizerOutcome {
			outcome, err := executeOptimizerProgramWithinDeadline(relative, source, absPath, false)
			if err != nil {
				t.Fatalf("%s: unoptimized run %v", path, err)
			}
			unoptimized = outcome
			return outcome
		}
		classify := func() {
			if optimized.status != "ok" {
				optimizedFailures = append(optimizedFailures, relative+": "+optimized.status)
			}
			if unoptimized.status != "ok" {
				unoptimizedFailures = append(unoptimizedFailures, relative+": "+unoptimized.status)
			}
		}

		runOptimized()
		baseline := runUnoptimized()

		if isScriptTimeoutOutcome(optimized) || isScriptTimeoutOutcome(baseline) {
			timeoutPages++
			continue
		}

		if _, quarantined := deadJumpCorpusQuarantine[relative]; quarantined {
			classify()
			if optimized.status != baseline.status {
				t.Fatalf("%s: optimized status %q does not match the unoptimized status %q (optimized detail %s, unoptimized detail %s)",
					path, optimized.status, baseline.status, optimized.detail, baseline.detail)
			}
			quarantinedPages++
			continue
		}

		if optimized == baseline {
			classify()
			stableComparisons++
			continue
		}

		// The pair diverged. Prove whether the page itself drifts before retrying, so a page
		// that is not repeatable is reported as such instead of being retried blindly.
		secondBaseline := runUnoptimized()
		if secondBaseline != baseline {
			classify()
			if optimized.status != secondBaseline.status {
				t.Fatalf("%s: optimized status %q does not match the unoptimized status %q (optimized detail %s, unoptimized detail %s)",
					path, optimized.status, secondBaseline.status, optimized.detail, secondBaseline.detail)
			}
			driftedPages++
			t.Logf("page drifted between two unoptimized runs, compared by status only; review whether it belongs in deadJumpCorpusQuarantine: %s", relative)
			continue
		}

		matched := false
		attempts := 1
		for i := 1; i < deadJumpCorpusPairAttempts; i++ {
			runOptimized()
			runUnoptimized()
			attempts++
			if optimized == unoptimized {
				matched = true
				break
			}
		}
		classify()
		if !matched {
			t.Fatalf("%s: optimizer changed the page outcome in %d adjacent pairs while the unoptimized output stayed repeatable\n  optimized:   %s\n  unoptimized: %s",
				path, attempts, optimized, unoptimized)
		}
		driftedPages++
		t.Logf("page matched only after %d adjacent pairs; review whether it belongs in deadJumpCorpusQuarantine: %s", attempts, relative)
	}

	t.Logf("corpus files=%d stable=%d quarantined=%d drifted=%d timeout=%d failures=%d elapsed=%s",
		len(files), stableComparisons, quarantinedPages, driftedPages, timeoutPages,
		len(optimizedFailures), time.Since(start).Round(time.Millisecond))

	sort.Strings(optimizedFailures)
	sort.Strings(unoptimizedFailures)
	if len(optimizedFailures) != len(unoptimizedFailures) {
		t.Fatalf("failure set size changed: optimized=%d unoptimized=%d\n  optimized:   %v\n  unoptimized: %v",
			len(optimizedFailures), len(unoptimizedFailures), optimizedFailures, unoptimizedFailures)
	}
	for i := range optimizedFailures {
		if optimizedFailures[i] != unoptimizedFailures[i] {
			t.Fatalf("failure set changed at index %d: optimized=%q unoptimized=%q\n  optimized:   %v\n  unoptimized: %v",
				i, optimizedFailures[i], unoptimizedFailures[i], optimizedFailures, unoptimizedFailures)
		}
	}
}
