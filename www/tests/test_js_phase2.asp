<%@ Language="JScript" %>
<%
/*
 * AxonASP JScript Phase 2 Test Page
 */

function assert(condition, message) {
    if (!condition) {
        Response.Write("FAIL: " + message + "<br>");
    } else {
        Response.Write("PASS: " + message + "<br>");
    }
}

// 1. Math Bindings
assert(Math.acosh(1) === 0, "Math.acosh(1) === 0");
assert(Math.sign(-5) === -1, "Math.sign(-5) === -1");
assert(Math.trunc(4.9) === 4, "Math.trunc(4.9) === 4");

// 2. Array Methods
var arr = [1, 2, 3];
arr.fill(4);
assert(arr.join(",") === "4,4,4", "Array.prototype.fill");

var arr2 = [1, 2, 3, 4, 5];
arr2.copyWithin(0, 3);
assert(arr2.join(",") === "4,5,3,4,5", "Array.prototype.copyWithin");

// 3. Object Symbols
var sym1 = Symbol("a");
var sym2 = Symbol("b");
var obj = { [sym1]: 1, [sym2]: 2, c: 3 };
var syms = Object.getOwnPropertySymbols(obj);
assert(syms.length === 2, "Object.getOwnPropertySymbols length");
assert(obj[syms[0]] === 1 || obj[syms[0]] === 2, "Retrieve via getOwnPropertySymbols");

Response.Write("Phase 2 Tests Completed.");
%>