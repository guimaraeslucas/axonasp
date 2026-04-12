<%
Option Explicit
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim o, n
Set o = New JSONobject
Set n = New JSONobject
n.Add "prop1", "x"
o.Add "objects", n
Response.Write "ISOBJ=" & IsObject(o("objects")) & "|TYPE=" & TypeName(o("objects")) & "|"
If IsObject(o("objects")) Then Response.Write o("objects").Serialize()
%>
