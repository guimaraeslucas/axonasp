<%
Response.Write "<h1>DocumentElement Test</h1>"

Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
xmlDoc.LoadXML("<root><item>Test</item></root>")

Response.Write "Step 1: Get DocumentElement<br>"
On Error Resume Next
Set docElem = xmlDoc.DocumentElement
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & " (Number: " & Err.Number & ")<br>"
    Err.Clear
Else
    Response.Write "No error<br>"
    If docElem Is Nothing Then
        Response.Write "DocumentElement is Nothing<br>"
    Else
        Response.Write "DocumentElement exists<br>"
        Response.Write "NodeName: " & docElem.NodeName & "<br>"
    End If
End If
On Error GoTo 0
%>
