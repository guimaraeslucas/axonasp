<%@ Language="JavaScript" %>
<%
function assert(condition, message) {
    if (!condition) {
        Response.Write("FAIL: " + message + "<br>");
    } else {
        Response.Write("PASS: " + message + "<br>");
    }
}

// Array.prototype.at
var arr = [10, 20, 30];
assert(arr.at(0) === 10, "arr.at(0)");
assert(arr.at(-1) === 30, "arr.at(-1)");
assert(arr.at(5) === undefined, "arr.at(5)");
assert("abc".at(1) === "b", "string.at(1)");

// Array.prototype.flat
var nested = [1, [2, [3]]];
assert(JSON.stringify(nested.flat()) === "[1,2,[3]]", "arr.flat()");
assert(JSON.stringify(nested.flat(2)) === "[1,2,3]", "arr.flat(2)");

// Array.prototype.flatMap
var fm = [1, 2].flatMap(x => [x, x * 10]);
assert(JSON.stringify(fm) === "[1,10,2,20]", "arr.flatMap()");

// Immutable methods
var original = [3, 1, 2];
var sorted = original.toSorted();
assert(JSON.stringify(original) === "[3,1,2]", "original unchanged after toSorted");
assert(JSON.stringify(sorted) === "[1,2,3]", "toSorted returned sorted array");

var reversed = original.toReversed();
assert(JSON.stringify(original) === "[3,1,2]", "original unchanged after toReversed");
assert(JSON.stringify(reversed) === "[2,1,3]", "toReversed returned reversed array");

var spliced = original.toSpliced(1, 1, 9);
assert(JSON.stringify(original) === "[3,1,2]", "original unchanged after toSpliced");
assert(JSON.stringify(spliced) === "[3,9,2]", "toSpliced returned modified array");

// Object.fromEntries
var entries = [["key1", "val1"], ["key2", "val2"]];
var obj = Object.fromEntries(entries);
assert(obj.key1 === "val1", "Object.fromEntries key1");
assert(obj.key2 === "val2", "Object.fromEntries key2");

// Object.values / Object.entries (verification)
var testObj = { a: 1, b: 2 };
assert(JSON.stringify(Object.values(testObj)) === "[1,2]", "Object.values");
assert(JSON.stringify(Object.entries(testObj)) === '[["a",1],["b",2]]', "Object.entries");

%>
