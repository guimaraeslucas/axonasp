<% @ Language = "VBScript" %>
<%
Dim notAnObj
notAnObj = 123

On Error Resume Next
notAnObj.X = 99

If Err.Number = 0 Or Not Instr(1, Err.Description, "required", 1) > 0 Then
    Response.Write "FAIL: Regression test did not raise Object Required error (Err.Number=" & Err.Number & ", Err.Description=" & Err.Description & ")"
    Response.End
End If

Response.Write "PASS: Regression behaviors are correct"
%>
