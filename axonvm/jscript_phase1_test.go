/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 */
package axonvm

import (
	"testing"
)

func TestJScriptArrayAt(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [1, 2, 3];
		Response.Write(arr.at(0) + ",");
		Response.Write(arr.at(1) + ",");
		Response.Write(arr.at(-1) + ",");
		Response.Write(arr.at(5) + ",");
		Response.Write("hello".at(0) + ",");
		Response.Write("hello".at(-1));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "1,2,3,undefined,h,o"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptArrayFlat(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [1, [2, [3, 4]]];
		Response.Write(JSON.stringify(arr.flat()) + "|");
		Response.Write(JSON.stringify(arr.flat(2)));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "[1,2,[3,4]]|[1,2,3,4]"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptArrayFlatMap(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [1, 2, 3];
		var res = arr.flatMap(x => [x, x * 2]);
		Response.Write(JSON.stringify(res));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "[1,2,2,4,3,6]"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptArrayImmutable(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [3, 1, 2];
		var sorted = arr.toSorted();
		var reversed = arr.toReversed();
		var spliced = arr.toSpliced(1, 1, 4, 5);
		Response.Write(JSON.stringify(arr) + "|");
		Response.Write(JSON.stringify(sorted) + "|");
		Response.Write(JSON.stringify(reversed) + "|");
		Response.Write(JSON.stringify(spliced));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "[3,1,2]|[1,2,3]|[2,1,3]|[3,4,5,2]"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptObjectFromEntries(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var entries = [["a", 1], ["b", 2]];
		var obj = Object.fromEntries(entries);
		Response.Write(obj.a + "," + obj.b);
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "1,2"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}
