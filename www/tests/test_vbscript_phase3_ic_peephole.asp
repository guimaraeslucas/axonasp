<%@ Language=VBScript %>
<%
Option Explicit

Class Probe
	Public Function Value()
		Value = 5
	End Function
End Class

Dim o, i, sum
Set o = New Probe
sum = 0

For i = 1 To 4
	sum = sum + o.Value()
Next

If False Then
	Response.Write "UNREACHABLE"
End If

Response.Write "SUM=" & sum
%>