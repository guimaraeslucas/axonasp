<%
Option Explicit
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim o
Set o = New JSONobject
o.Add "nome", "abc"
Response.Write "D1=" & o.value("nome") & "|D2=" & o("nome")
%>
