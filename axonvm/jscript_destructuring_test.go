/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 */
package axonvm

import (
	"testing"
)

func TestJScriptDestructuringObject(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var obj = { a: 10, b: 20 };
		var { a, b } = obj;
		Response.Write(a + "|" + b);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "10|20" {
		t.Errorf("expected '10|20', got %q", out)
	}
}

func TestJScriptDestructuringObjectNested(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var obj = { a: { x: 100 }, b: 200 };
		var { a: { x }, b } = obj;
		Response.Write(x + "|" + b);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "100|200" {
		t.Errorf("expected '100|200', got %q", out)
	}
}

func TestJScriptDestructuringObjectComputed(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var k = "prop";
		var obj = { prop: "hello" };
		var { [k]: val } = obj;
		Response.Write(val);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "hello" {
		t.Errorf("expected 'hello', got %q", out)
	}
}

func TestJScriptDestructuringObjectAssignment(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var a, b;
		({ a, b } = { a: 1, b: 2 });
		Response.Write(a + "|" + b);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "1|2" {
		t.Errorf("expected '1|2', got %q", out)
	}
}

func TestJScriptDestructuringObjectNull(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		try {
			var { a } = null;
		} catch (e) {
			Response.Write(e.indexOf("TypeError") !== -1);
		}
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "True" {
		t.Errorf("expected 'True', got %q", out)
	}
}

func TestJScriptDestructuringArray(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var [a, b, c] = [1, 2, 3];
		Response.Write(a + "|" + b + "|" + c);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "1|2|3" {
		t.Errorf("expected '1|2|3', got %q", out)
	}
}

func TestJScriptDestructuringArrayNested(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var [a, [b, c]] = [10, [20, 30]];
		Response.Write(a + "|" + b + "|" + c);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "10|20|30" {
		t.Errorf("expected '10|20|30', got %q", out)
	}
}

func TestJScriptDestructuringArrayString(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var [x, y, z] = "ABC";
		Response.Write(x + "|" + y + "|" + z);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "A|B|C" {
		t.Errorf("expected 'A|B|C', got %q", out)
	}
}

func TestJScriptDestructuringArrayElision(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var [, , x] = [1, 2, 3];
		Response.Write(x);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "3" {
		t.Errorf("expected '3', got %q", out)
	}
}

func TestJScriptDestructuringArrayNonIterable(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		try {
			var [a] = 123;
		} catch (e) {
			Response.Write(e.indexOf("not iterable") !== -1);
		}
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "True" {
		t.Errorf("expected 'True', got %q", out)
	}
}

func TestJScriptDestructuringDefaultValues(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var { a = 10, b = 20 } = { a: 1 };
		var [x = 100, y = 200] = [1];
		Response.Write(a + "|" + b + "|" + x + "|" + y);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "1|20|1|200" {
		t.Errorf("expected '1|20|1|200', got %q", out)
	}
}

func TestJScriptDestructuringRest(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var { a, ...rest } = { a: 1, b: 2, c: 3 };
		var [x, ...arrayRest] = [10, 20, 30];
		Response.Write(a + "|" + JSON.stringify(rest) + "|" + x + "|" + JSON.stringify(arrayRest));
	`))
	if err != nil {
		t.Fatal(err)
	}
	// Object rest: JSON.stringify might vary key order, but here it's likely {"b":2,"c":3}
	// Array rest: [20, 30]
	if out != "1|{\"b\":2,\"c\":3}|10|[20,30]" {
		t.Errorf("expected '1|{\"b\":2,\"c\":3}|10|[20,30]', got %q", out)
	}
}

func TestJScriptDestructuringDeeplyNested(t *testing.T) {
	// Sub-phase 5.5: Deeply nested (depth 10)
	out, err := runJScript2(t, jscriptSrc(`
		var nested = { a: { a: { a: { a: { a: { a: { a: { a: { a: { a: 42 } } } } } } } } } };
		var { a: { a: { a: { a: { a: { a: { a: { a: { a: { a: val } } } } } } } } } } = nested;
		Response.Write(val);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "42" {
		t.Errorf("expected '42', got %q", out)
	}
}
