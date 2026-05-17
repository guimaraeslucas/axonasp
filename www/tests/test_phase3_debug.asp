<%@ Language="JScript" %>
<%
var obj = { x: 1, y: 2 };
obj[Symbol.unscopables] = { x: true };
var x = 10, y = 20;
with (obj) {
    Response.Write("with x = " + x + "<br>");
    Response.Write("with y = " + y + "<br>");
}
%>