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
