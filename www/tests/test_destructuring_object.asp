<%@ Language="JScript" %>
<%
/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 */
    Response.Write("Testing Object Destructuring (Sub-Phase 5.2)\n");

    // 1. Simple var destructuring
    var obj = { x: 10, y: 20 };
    var { x, y } = obj;
    Response.Write("x: " + x + ", y: " + y + "\n");

    // 2. let destructuring with block scope
    {
        let { x: x2, y: y2 } = { x: 100, y: 200 };
        Response.Write("x2: " + x2 + ", y2: " + y2 + "\n");
    }
    // x2 should be undefined here
    Response.Write("typeof x2: " + typeof x2 + "\n");

    // 3. Nested destructuring
    var nested = { a: { b: { c: "nested value" } } };
    var { a: { b: { c } } } = nested;
    Response.Write("c: " + c + "\n");

    // 4. Computed properties
    var key = "dynamic";
    var { [key]: val } = { dynamic: "hello computed" };
    Response.Write("val: " + val + "\n");

    // 5. Assignment destructuring
    var a, b;
    ({ a, b } = { a: "A", b: "B" });
    Response.Write("a: " + a + ", b: " + b + "\n");

    // 6. Null/Undefined check
    try {
        var { p } = null;
    } catch (e) {
        Response.Write("Caught error (null): " + (e.indexOf("TypeError") !== -1) + "\n");
    }

    try {
        var { q } = undefined;
    } catch (e) {
        Response.Write("Caught error (undefined): " + (e.indexOf("TypeError") !== -1) + "\n");
    }
%>
