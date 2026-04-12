<%
@ Language = VBScript
%>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8" />
        <title>Bridge Debug Test</title>
    </head>
    <body>
        <h1>Bridge Debug Test</h1>
        <pre>
<%
Response.Write "Test 1: Simple string function" & vbCrLf
Response.Write "AxHostNameValue(): " & AxHostNameValue() & vbCrLf

Response.Write vbCrLf & "Test 2: Get environment list" & vbCrLf
Dim envItems
envItems = AxEnvironmentList()
Response.Write "Got envItems: " & TypeName(envItems) & vbCrLf

Response.Write vbCrLf & "Test 3: Try UBound" & vbCrLf
On Error Resume Next
Dim ub
ub = UBound(envItems)
If Err.Number <> 0 Then
    Response.Write "ERROR in UBound: " & Err.Description & vbCrLf
Else
    Response.Write "UBound(envItems) = " & ub & vbCrLf
End If
On Error Goto 0

Response.Write vbCrLf & "Test 4: Try subscript access" & vbCrLf
On Error Resume Next
If Err.Number = 0 Then
    Response.Write "envItems(0) = " & envItems(0) & vbCrLf
End If
If Err.Number <> 0 Then
    Response.Write "ERROR in subscript: " & Err.Description & vbCrLf
End If
On Error Goto 0

Response.Write vbCrLf & "Done!" & vbCrLf
%>
</pre>
    </body>
</html>
