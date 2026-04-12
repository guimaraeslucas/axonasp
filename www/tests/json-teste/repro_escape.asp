<%
Option Explicit
On Error Resume Next
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim o
Set o = New JSONobject
Response.Write o.EscapeCharacters("abc")
If Err.Number <> 0 Then Response.Write "|ERR:" & Err.Description
%>
