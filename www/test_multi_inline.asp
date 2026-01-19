<%@ language="VBScript" %>
<% Option Explicit %>
<% Dim value: value = 5
If value > 3 Then value = value + 1 : value = value * 2 : End If
Response.Write value %>