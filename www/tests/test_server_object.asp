<%
' Test if Server object is accessible
Response.Write("Starting test...<br>")
Response.Write("Server object type: " & VarType(Server) & "<br>")

If IsEmpty(Server) Then
    Response.Write("ERROR: Server is Empty!<br>")
ElseIf IsNull(Server) Then
    Response.Write("ERROR: Server is Null!<br>")
Else
    Response.Write("Server object is valid<br>")
    Response.Write("Attempting CreateObject...<br>")
    
    On Error Resume Next
    Err.Clear
    
    Set dict = Server.CreateObject("Scripting.Dictionary")
    
    If Err.Number <> 0 Then
        Response.Write("ERROR: " & Err.Number & " - " & Err.Description & "<br>")
    Else
        Response.Write("SUCCESS: Dictionary created<br>")
        dict.Add "test", "value"
        Response.Write("Added item to dictionary<br>")
        Response.Write("Dictionary.Count = " & dict.Count & "<br>")
    End If
End If

Response.Write("Test complete<br>")
%>
