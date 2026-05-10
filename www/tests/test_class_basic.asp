<%@ Language="JavaScript" %>
<%
function assert(condition, message) {
    if (!condition) {
        Response.Write("Assertion failed: " + message + "<br>");
    }
}

// 1. Basic Class Declaration and Instantiation
class A {
    constructor() {
        this.x = 42;
    }
}

var a = new A();
assert(a instanceof A, "a should be an instance of A");
assert(a.x === 42, "a.x should be 42");

// 2. Class Expression
var B = class {
    constructor(v) {
        this.v = v;
    }
};

var b = new B(100);
assert(b.v === 100, "b.v should be 100");

// 3. Class Constructor enforcement (must use 'new')
try {
    A();
    Response.Write("Expected error but none thrown (A())<br>");
} catch (e) {
    var msg = "" + e;
    if (msg.indexOf("Class constructor cannot be invoked without 'new'") === -1) {
        Response.Write("Expected error message containing 'new', got: '" + msg + "'<br>");
    } else {
        // Success
    }
}

// 4. TDZ for classes
try {
    var c = new C();
    class C {}
    Response.Write("Expected error but none thrown (TDZ C)<br>");
} catch (e) {
    var msg = "" + e;
    if (msg.indexOf("before initialization") === -1) {
        Response.Write("Expected TDZ error, got: " + msg + "<br>");
    }
}

// 5. Default constructor
class D {}
var d = new D();
assert(d instanceof D, "d should be an instance of D");

Response.Write("DONE");
%>
