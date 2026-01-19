<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit

Dim x
x = 5

If x = 5 Then x = x + 1 : Response.Write "test"

Response.Write "x = " & x
%>
