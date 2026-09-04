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
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package axonvm

import (
	"strings"
	"testing"
)

// TestJScriptES5PrototypeNewTopLevelMemberAssign is a regression test for the
// JScript "Stack underflow" (800A0033) crash reported on ES5 prototype chains
// combined with the `new` operator.
//
// Root cause: commit aad6024 removed the trailing OpJSLoadUndefined that kept
// member/index assignments in balance. OpJSMemberSet and OpJSIndexSet consume
// the assigned value and push nothing, yet every expression statement still
// emits OpJSPop. At the root frame (no local-slot headroom below sp) that pop
// drove sp below zero and raised "Stack underflow". The compiler now retains a
// copy of the assigned value, matching JavaScript assignment-expression
// semantics (`obj.prop = value` evaluates to `value`).
func TestJScriptES5PrototypeNewTopLevelMemberAssign(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function Box(v) { this.value = v; }` +
		`Box.prototype.getValue = function() { return this.value; };` +
		`Array.prototype.firstItem = function() { return this[0]; };` +
		`var box = new Box(11);` +
		`var values = [42, 99];` +
		`Response.Write("BOX=" + box.getValue() + ";");` +
		`Response.Write("ARR=" + values.firstItem() + ";");` +
		`Response.Write("PROTO=" + Object.getPrototypeOf(box).constructor.prototype.getValue.call(box) + ";");` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "BOX=11;ARR=42;PROTO=11;" {
		t.Fatalf("unexpected ES5 prototype/new output: %q", out)
	}
}

// TestJScriptMemberAssignStatementStackParity exercises every assignment form
// that previously underflowed when used as an expression statement: dot
// member-set, bracket index-set, private dot-set (via classes), chained
// member-set on the constructor `this`, and prototype extension. All run at
// the top-level root frame where a missing expression result used to drive sp
// below zero.
func TestJScriptMemberAssignStatementStackParity(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var o = { count: 0 };` +
		`o.count = o.count + 1;` +
		`o["count"] = o.count + 1;` +
		`var arr = [1, 2, 3];` +
		`arr[1] = 20;` +
		`String.prototype.repeat2 = function() { return this + this; };` +
		`var s = "ab";` +
		`function Ctor(v) { this.value = v; }` +
		`var inst = new Ctor(5);` +
		`inst.value = inst.value + 1;` +
		`inst["value"] = inst.value + 1;` +
		`var nested = { inner: { deep: 0 } };` +
		`nested.inner.deep = 7;` +
		`Response.Write(o.count + "|" + arr[1] + "|" + "ab".repeat2() + "|" + inst.value + "|" + nested.inner.deep);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "2|20|abab|7|7" {
		t.Fatalf("unexpected member-assign statement parity output: %q", out)
	}
}

// TestJScriptMemberAssignmentExpressionValue verifies the retained assignment
// result is the assigned (right-hand side) value, matching ECMAScript
// assignment-expression semantics. Before the fix the engine left no value (or,
// pre-regression, `undefined`), so nested/chained assignments like
// `a = b.x = 7` could not propagate the assigned value.
func TestJScriptMemberAssignmentExpressionValue(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var o = {};` +
		`var out = (o.x = 42);` +
		`var a; var b = { y: 0 };` +
		`a = b.y = 7;` +
		`var arr = [1];` +
		`var idx = (arr[0] = 99);` +
		`var cond = ((o.z = 3) === 3) ? "T" : "F";` +
		`Response.Write(out + "|" + o.x + "|" + a + "|" + b.y + "|" + idx + "|" + arr[0] + "|" + cond);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "42|42|7|7|99|99|T" {
		t.Fatalf("unexpected member-assign expression value output: %q", out)
	}
}

// TestJScriptRuntimeErrorSourceIsJScriptNotVBScript verifies that uncaught
// JScript runtime errors are surfaced to the host with Source
// "JScript runtime error" and Category "JScript runtime". The shared dispatch
// loop must route stack/execution faults raised while JScript state is active
// through the JScript runtime error path so the host never mislabels them as
// VBScript runtime errors (see the error routing directive).
func TestJScriptRuntimeErrorSourceIsJScriptNotVBScript(t *testing.T) {
	// Calling a name that is never defined throws an uncaught JScript runtime
	// error ("Undefined identifier" 5009 or "Function expected" 5002).
	source := `<script runat="server" language="JScript">` +
		`definitelyNotDefinedAnywhere(1);` +
		`</script>`

	_, err := runASPSourceForTestWithErr(t, source)
	if err == nil {
		t.Fatalf("expected a JScript runtime error")
	}

	vme, ok := err.(*VMError)
	if !ok {
		t.Fatalf("expected *VMError, got %T: %v", err, err)
	}
	if vme.Source != "JScript runtime error" {
		t.Fatalf("expected Source %q, got %q", "JScript runtime error", vme.Source)
	}
	if vme.Category != "JScript runtime" {
		t.Fatalf("expected Category %q, got %q", "JScript runtime", vme.Category)
	}
	if strings.Contains(vme.Source, "VBScript") || strings.Contains(vme.Category, "VBScript") {
		t.Fatalf("JScript runtime error must not be labeled as VBScript: source=%q category=%q", vme.Source, vme.Category)
	}
	if vme.Description == "" {
		t.Fatalf("expected a non-empty runtime error description")
	}
}
