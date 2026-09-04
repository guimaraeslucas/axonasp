package axonvm

import (
	"testing"
	"time"
)

// TestJScriptNestedLoopDoesNotCorruptCallerLocals guards the JScript stack
// integrity regression where a callee running high-iteration loops (fast-var
// for-loops, while loops, and standalone local ++/-- statements) wrote through
// the caller's stack frame, corrupting its locals (observed as the benchmark
// logging "undefined completed in 1788315467637 ms").
//
// Before the fix, update expressions on local slots (e.g. `count++` used as an
// expression statement) consumed an extra stack slot per iteration, and loops
// inside nested function calls corrupted the caller's `name`, `start` and `end`
// locals. This test pins correct values and correct loop results.
func TestJScriptNestedLoopDoesNotCorruptCallerLocals(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function bench(n) {` +
		`  var count = 0;` +
		`  var it = n;` +
		`  for (var i = 0; i < it; i++) { count++; }` +
		`  var j = it;` +
		`  while (j > 0) { j--; count--; }` +
		`  return count;` +
		`}` +
		`function run(name, fn, n) {` +
		`  var start = 1111;` +
		`  var r = fn(n);` +
		`  var end = 2222;` +
		`  return name + "|" + start + "|" + end + "|" + r;` +
		`}` +
		`Response.Write(run("alpha", bench, 1000000));` +
		`</script>`

	done := make(chan string, 1)
	go func() { done <- runASPSourceForTest(t, source) }()
	select {
	case out := <-done:
		expected := "alpha|1111|2222|0"
		if out != expected {
			t.Fatalf("caller locals or loop result corrupted: got %q, want %q", out, expected)
		}
	case <-time.After(30 * time.Second):
		t.Fatal("nested loop benchmark did not terminate (possible runaway/hang)")
	}
}

// TestJScriptObjectKeyTrackingManyKeysCompletes guards the O(N^2) regression in
// jsTrackObjectKey: tracking insertion order by linearly rescanning the whole
// key-order slice on every new property made building a single large object
// quadratic (100k distinct keys took ~24s and 1M keys hung the server). The
// insertion-order index must be O(1), so building and reading a large object
// completes quickly with correct values.
func TestJScriptObjectKeyTrackingManyKeysCompletes(t *testing.T) {
	if testing.Short() {
		t.Skip("large-object regression test")
	}
	source := `<script runat="server" language="JScript">` +
		`var obj = {};` +
		`var n = 150000;` +
		`for (var i = 0; i < n; i++) { obj["key_" + i] = i; }` +
		`var check = true;` +
		`for (var j = 0; j < n; j++) { if (obj["key_" + j] !== j) { check = false; } }` +
		`var keys = 0;` +
		`for (var k in obj) { keys++; }` +
		`Response.Write((check ? "ok" : "bad") + "|" + keys);` +
		`</script>`

	done := make(chan string, 1)
	start := time.Now()
	go func() { done <- runASPSourceForTest(t, source) }()
	select {
	case out := <-done:
		expected := "ok|150000"
		if out != expected {
			t.Fatalf("large-object build/read produced wrong result: got %q, want %q", out, expected)
		}
		elapsed := time.Since(start)
		if elapsed > 15*time.Second {
			t.Fatalf("large-object build took too long (%v); quadratic key tracking likely regressed", elapsed.Round(time.Millisecond))
		}
	case <-time.After(20 * time.Second):
		t.Fatal("large-object build hung (quadratic key tracking regression)")
	}
}

// TestJScriptExpressionStatementUpdateValueRetention verifies that update
// expressions (`c++`, `++c`) used as expression statements and as function
// return values keep the operand stack balanced (no net underflow/overflow per
// statement) while producing correct prefix/postfix semantics.
func TestJScriptExpressionStatementUpdateValueRetention(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function f() {` +
		`  var c = 0;` +
		`  c++;` +
		`  c++;` +
		`  ++c;` +
		`  var pre = ++c;` +
		`  var post = c++;` +
		`  return c + "|" + pre + "|" + post;` +
		`}` +
		`Response.Write(f());` +
		`</script>`

	done := make(chan string, 1)
	go func() { done <- runASPSourceForTest(t, source) }()
	select {
	case out := <-done:
		// c starts 0: c++ ->1, c++ ->2, ++c ->3, then pre=++c => c=4, pre=4,
		// then post=c++ => post=4, c=5.
		want := "5|4|4"
		if out != want {
			t.Fatalf("update expression semantics wrong: got %q, want %q", out, want)
		}
	case <-time.After(20 * time.Second):
		t.Fatal("update-expression test hung (stack imbalance likely)")
	}
}
