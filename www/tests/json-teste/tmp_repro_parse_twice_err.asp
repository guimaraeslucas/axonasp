<%
Option Explicit
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, s
Set jsonObj = New JSONobject
s = "[{""a"":1}]"
On Error Resume Next
jsonObj.parse s
Response.Write "E1=" & Err.Number & " " & Err.Description & "<br>"
Err.Clear
jsonObj.parse s
Response.Write "E2=" & Err.Number & " " & Err.Description & "<br>"
Err.Clear
On Error Goto 0
Response.Write "done"
%>
