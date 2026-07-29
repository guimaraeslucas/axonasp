<%@ Language=VBScript %>
<%
Response.Write("START" & vbCrLf)
On Error Resume Next
Dim x
x = 1 / 0
If Err.Number <> 0 Then
    Response.Write("ERROR CAUGHT: " & Err.Description & vbCrLf)
Else
    Response.Write("NO ERROR" & vbCrLf)
End If
On Error GoTo 0
Response.Write("END" & vbCrLf)
%>