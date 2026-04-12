<!--#include file="jsonObject.class.asp" -->
<%
Option Explicit
Dim j
Set j = New JSONobject
On Error Resume Next
j.Add "x", Now
Response.Write "Err=" & Err.Number & "|" & Err.Description & "|Type=" & TypeName(j.value("x"))
%>
