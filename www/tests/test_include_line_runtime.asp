<%@ LANGUAGE=VBScript %>
<!--#include file="test_include_line_big.inc"-->
<%
On Error Resume Next

Dim a
a = 10
Dim b
b = 0
Dim c
c = a / b

If Err.Number <> 0 Then
    Response.Write "runtime-line=" & Err.Line & "\n"
    Response.Write "runtime-number=" & Err.Number & "\n"
Else
    Response.Write "no-error"
End If
%>
