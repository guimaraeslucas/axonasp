<!--#include file="jsonObject.class.asp" -->
<%
Option Explicit
Dim p
Set p = New JSONpair
On Error Resume Next
p.value = Now
Response.Write "AfterLet|Err=" & Err.Number & "|" & Err.Description & "|Type=" & TypeName(p.value)
%>
