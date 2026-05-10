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
    Response.Write("Testing Iteration Protocol (Sub-Phase 5.1)\n");

    // Array Iterator
    var arr = [1, 2, 3];
    var it = arr[Symbol.iterator]();
    Response.Write("Array it.next().value (1): " + it.next().value + "\n");
    Response.Write("Array it.next().value (2): " + it.next().value + "\n");
    Response.Write("Array it.next().value (3): " + it.next().value + "\n");
    Response.Write("Array it.next().done: " + it.next().done + "\n");

    // String Iterator
    var s = "XYZ";
    var sit = s[Symbol.iterator]();
    Response.Write("String it.next().value (X): " + sit.next().value + "\n");
    Response.Write("String it.next().value (Y): " + sit.next().value + "\n");
    Response.Write("String it.next().value (Z): " + sit.next().value + "\n");
    Response.Write("String it.next().done: " + sit.next().done + "\n");

    // for...of with custom iterable
    var myIterable = {
        [Symbol.iterator]: function() {
            var i = 0;
            return {
                next: function() {
                    i++;
                    return { value: i * 10, done: i > 3 };
                }
            };
        }
    };
    Response.Write("Custom iterable for...of: ");
    for (var x of myIterable) {
        Response.Write(x + " ");
    }
    Response.Write("\n");
    
    // built-in for...of verification
    Response.Write("Array for...of: ");
    for (var v of [10, 20, 30]) {
        Response.Write(v + " ");
    }
    Response.Write("\n");
%>
