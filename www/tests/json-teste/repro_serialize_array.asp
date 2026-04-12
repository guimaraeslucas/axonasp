<%
Option Explicit
On Error Resume Next
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim o, arr
Set o = New JSONobject
arr = Array(1,"2",3)
Response.Write o.serializeArray(arr)
If Err.Number <> 0 Then Response.Write "|ERR:" & Err.Description
%>
