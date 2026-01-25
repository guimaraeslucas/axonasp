<%
Response.Write "<h1>GetElementsByTagName Test</h1>"

Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
xmlDoc.LoadXML("<root><item>First</item><item>Second</item><item>Third</item></root>")

Response.Write "Step 1: Call GetElementsByTagName<br>"
On Error Resume Next
Set items = xmlDoc.GetElementsByTagName("item")
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & " (Number: " & Err.Number & ")<br>"
    Err.Clear
Else
    Response.Write "No error calling GetElementsByTagName<br>"
    If items Is Nothing Then
        Response.Write "Result is Nothing<br>"
    Else
        Response.Write "Result is not Nothing<br>"
        Response.Write "TypeName: " & TypeName(items) & "<br>"
        
        ' Try to get UBound
        ub = UBound(items)
        If Err.Number <> 0 Then
            Response.Write "ERROR getting UBound: " & Err.Description & "<br>"
            Err.Clear
        Else
            Response.Write "UBound: " & ub & "<br>"
            Response.Write "Count: " & (ub + 1) & "<br>"
            
            ' Try to access elements
            For i = 0 To ub
                Set elem = items(i)
                If elem Is Nothing Then
                    Response.Write "  Item " & i & " is Nothing<br>"
                Else
                    Response.Write "  Item " & i & ": " & elem.NodeName & " = " & elem.Text & "<br>"
                End If
            Next
        End If
    End If
End If
On Error GoTo 0
%>
