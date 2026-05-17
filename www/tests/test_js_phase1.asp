<%@ Language="JScript" %>
<%
/*
 * AxonASP JScript Phase 1 Test Page
 */

function assert(condition, message) {
    if (!condition) {
        Response.Write("FAIL: " + message + "<br>");
    } else {
        Response.Write("PASS: " + message + "<br>");
    }
}

// 1. Numeric Literals
assert(0b1010 === 10, "Binary literal 0b1010 === 10");
assert(0o744 === 484, "Octal literal 0o744 === 484");

// 2. Object Shorthand
var x = 10, y = 20;
var obj = { x: x, y: y };
var obj2 = { x, y };
assert(obj2.x === 10 && obj2.y === 20, "Object shorthand {x, y}");

// 3. Method Shorthand
var obj3 = {
    val: 42,
    getVal() { return this.val; }
};
assert(obj3.getVal() === 42, "Method shorthand getVal() {}");

// 4. new.target
function Foo() {
    this.newTarget = new.target;
}
var f1 = new Foo();
assert(f1.newTarget === Foo, "new.target in constructor call");
Foo();
// can't easily test Foo() call result here without returning it, 
// but we verified it in Go tests.

// 5. Computed Properties
var key = "dynamicKey";
var obj4 = {
    [key]: "dynamicValue",
    ["prop" + (1+2)]: 123
};
assert(obj4.dynamicKey === "dynamicValue", "Computed property [key]");
assert(obj4.prop3 === 123, "Computed property ['prop'+3]");

// Symbol keys
var sym = Symbol("test");
var obj5 = {
    [sym]: "symVal"
};
assert(obj5[sym] === "symVal", "Computed property with Symbol key");

Response.Write("Phase 1 Tests Completed.");
%>
