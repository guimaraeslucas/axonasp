<%@ language="JScript" %>
<%
    var wm = new WeakMap();
    var key = Symbol("test");
    wm.set(key, "WeakMap value");
    Response.Write("WeakMap has: " + wm.has(key) + "<br>");
    Response.Write("WeakMap get: " + wm.get(key) + "<br>");
    
    var ws = new WeakSet();
    ws.add(key);
    Response.Write("WeakSet has: " + ws.has(key) + "<br>");
    
    // Negative test: registered symbols should throw
    try {
        wm.set(Symbol.for("registered"), "fail");
        Response.Write("FAIL: Registered symbol allowed in WeakMap<br>");
    } catch(e) {
        Response.Write("PASS: Registered symbol threw: " + e.message + "<br>");
    }
    
    // Negative test: well-known symbols should throw
    try {
        wm.set(Symbol.iterator, "fail");
        Response.Write("FAIL: Well-known symbol allowed in WeakMap<br>");
    } catch(e) {
        Response.Write("PASS: Well-known symbol threw: " + e.message + "<br>");
    }
%>
