<%
Option Explicit
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim o, arr
Set o = New JSONobject
arr = Array(1,"2",3)
Response.Write o.serializeArray(arr)
Response.Write "|done"
%>
