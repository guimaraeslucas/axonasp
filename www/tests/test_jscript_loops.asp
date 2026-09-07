<%@ Language="JScript" %>
<%
// Test JScript loop compilation and execution (single-language JScript page).
// NOTE: This page intentionally avoids mixing JScript <script runat="server">
// blocks with a VBScript body. In this engine (as in standard ASP) server-side
// script blocks of a different language are hoisted to run before the VBScript
// body, and cross-language page globals are not shared, so a VBScript-created
// object cannot be fed into the JScript blocks and read back later.
var t = Server.CreateObject("G3TESTSUITE");

// Test 1: While loop
t.Describe("JScript While Loop");
var i = 0;
var result = "";
while (i < 3) {
    result += i + ",";
    i++;
}
t.AssertEquals("0,1,2,", result, "While loop should iterate 0,1,2");

// Test 2: Do-While loop
t.Describe("JScript Do-While Loop");
var j = 0;
result = "";
do {
    result += j + ",";
    j++;
} while (j < 2);
t.AssertEquals("0,1,", result, "Do-while loop should iterate 0,1");

// Test 3: For loop
t.Describe("JScript For Loop");
result = "";
for (var k = 0; k < 3; k++) {
    result += k + ",";
}
t.AssertEquals("0,1,2,", result, "For loop should iterate 0,1,2");

// Test 4: Break statement
t.Describe("JScript Break Statement");
result = "";
for (var m = 0; m < 5; m++) {
    if (m === 3) break;
    result += m + ",";
}
t.AssertEquals("0,1,2,", result, "Break should exit at m=3");

// Test 5: Continue statement
t.Describe("JScript Continue Statement");
result = "";
for (var n = 0; n < 5; n++) {
    if (n === 2) continue;
    result += n + ",";
}
t.AssertEquals("0,1,3,4,", result, "Continue should skip n=2");

// Test 6: Arithmetic operators
t.Describe("JScript Arithmetic Operators");
t.AssertEquals(
    "3,12,5,1",
    (5 - 2) + "," + (3 * 4) + "," + (10 / 2) + "," + (10 % 3),
    "Arithmetic operators should work correctly"
);

// Test 7: Comparison operators
t.Describe("JScript Comparison Operators");
t.AssertEquals(
    "true,true,true,false",
    (5 > 3 ? "true" : "false") + "," + (2 < 3 ? "true" : "false") + "," + (5 >= 5 ? "true" : "false") + "," + (3 <= 2 ? "true" : "false"),
    "Comparison operators should work correctly"
);

// Test 8: Logical operators
t.Describe("JScript Logical Operators");
t.AssertEquals(
    "true,true,false",
    (true && true ? "true" : "false") + "," + (false || true ? "true" : "false") + "," + (!true ? "true" : "false"),
    "Logical operators should work correctly"
);

t.Summary();
%>