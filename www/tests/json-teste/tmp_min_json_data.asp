<!--#include file="jsonObject.class.asp" -->
<%
Option Explicit
Dim o
Set o = New JSONobject
Set o = o.parse("[{""a"":1}]")
Response.Write "TypeName(o)=" & TypeName(o) & "<br>"
On Error Resume Next
o.Add "data", Now
Response.Write "Err=" & Err.Number & "|" & Err.Description & "<br>"
%>
