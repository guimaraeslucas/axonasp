<%
debug_asp_code = "TRUE"
Response.Write "<h1>MSXML Debug Test</h1>"

Response.Write "<h3>Test: Create DOMDocument</h3>"
On Error Resume Next
Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
If Err.Number <> 0 Then
    Response.Write "ERROR creating DOMDocument: " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "DOMDocument created successfully<br>"
    Response.Write "TypeName: " & TypeName(xmlDoc) & "<br>"
End If

Response.Write "<h3>Test: LoadXML</h3>"
simpleXML = "<root><item>Test</item></root>"
result = xmlDoc.LoadXML(simpleXML)
If Err.Number <> 0 Then
    Response.Write "ERROR in LoadXML: " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "LoadXML returned: " & result & "<br>"
End If

Response.Write "<h3>Test: DocumentElement Property</h3>"
Set docElem = xmlDoc.DocumentElement
If Err.Number <> 0 Then
    Response.Write "ERROR accessing DocumentElement: " & Err.Description & "<br>"
    Err.Clear
ElseIf docElem Is Nothing Then
    Response.Write "DocumentElement is Nothing<br>"
Else
    Response.Write "DocumentElement exists: " & docElem.NodeName & "<br>"
End If

Response.Write "<h3>Test: GetElementsByTagName</h3>"
Set items = xmlDoc.GetElementsByTagName("item")
If Err.Number <> 0 Then
    Response.Write "ERROR in GetElementsByTagName: " & Err.Description & "<br>"
    Err.Clear
ElseIf items Is Nothing Then
    Response.Write "GetElementsByTagName returned Nothing<br>"
Else
    Response.Write "GetElementsByTagName returned an object<br>"
    Response.Write "TypeName: " & TypeName(items) & "<br>"
    On Error Resume Next
    ub = UBound(items)
    If Err.Number <> 0 Then
        Response.Write "ERROR getting UBound: " & Err.Description & "<br>"
        Err.Clear
    Else
        Response.Write "UBound: " & ub & "<br>"
        Response.Write "Found " & (ub + 1) & " items<br>"
    End If
End If

Response.Write "<h3>Test: CreateElement</h3>"
Set newElem = xmlDoc.CreateElement("test")
If Err.Number <> 0 Then
    Response.Write "ERROR in CreateElement: " & Err.Description & "<br>"
    Err.Clear
ElseIf newElem Is Nothing Then
    Response.Write "CreateElement returned Nothing<br>"
Else
    Response.Write "CreateElement returned: " & newElem.NodeName & "<br>"
End If

On Error GoTo 0
%>
