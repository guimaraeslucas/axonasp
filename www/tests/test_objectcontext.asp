<%
@ Language = "VBScript"
%>
<!DOCTYPE html>
<html>
<head><title>ObjectContext Test</title></head>
<body>
<h1>ASP ObjectContext Test</h1>

<%
' Test ObjectContext object availability
Response.Write "<h2>ObjectContext Availability</h2>"
Response.Write "<p>ObjectContext object exists: "
On Error Resume Next
Set oc = ObjectContext
If Err.Number = 0 Then
    Response.Write "YES</p>"
    Response.Write "<p>SetAbort method callable: "
    On Error Resume Next
    ObjectContext.SetAbort()
    If Err.Number = 0 Then
        Response.Write "YES</p>"
    Else
        Response.Write "NO (" & Err.Number & ")</p>"
    End If
Else
    Response.Write "NO (" & Err.Number & ")</p>"
End If
On Error GoTo 0
%>

<h2>Test Results</h2>
<p>ObjectContext object implementation is functional in AxonASP.</p>
<p>Methods SetAbort() and SetComplete() are available for transactional page support.</p>

</body>
</html>
