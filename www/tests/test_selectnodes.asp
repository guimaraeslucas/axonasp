<%
Response.Write "<h1>CreateElement and SelectSingleNode Test</h1>"

Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
xmlDoc.LoadXML("<root><users><user><name>Alice</name><id>1</id></user></users></root>")

Response.Write "<h3>Test 1: CreateElement</h3>"
On Error Resume Next
Set newElem = xmlDoc.CreateElement("newtag")
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & "<br>"
    Err.Clear
ElseIf newElem Is Nothing Then
    Response.Write "CreateElement returned Nothing<br>"
Else
    Response.Write "CreateElement succeeded<br>"
    Response.Write "NodeName: " & newElem.NodeName & "<br>"
End If

Response.Write "<h3>Test 2: SelectSingleNode</h3>"
Set nameNode = xmlDoc.SelectSingleNode("//name")
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & "<br>"
    Err.Clear
ElseIf nameNode Is Nothing Then
    Response.Write "SelectSingleNode returned Nothing<br>"
Else
    Response.Write "SelectSingleNode succeeded<br>"
    Response.Write "NodeName: " & nameNode.NodeName & "<br>"
    Response.Write "Text: " & nameNode.Text & "<br>"
End If

Response.Write "<h3>Test 3: SelectNodes</h3>"
Set nodes = xmlDoc.SelectNodes("//user")
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & "<br>"
    Err.Clear
ElseIf nodes Is Nothing Then
    Response.Write "SelectNodes returned Nothing<br>"
Else
    Response.Write "SelectNodes succeeded<br>"
    ub = UBound(nodes)
    If Err.Number <> 0 Then
        Response.Write "ERROR getting UBound: " & Err.Description & "<br>"
        Err.Clear
    Else
        Response.Write "Found " & (ub + 1) & " nodes<br>"
    End If
End If

On Error GoTo 0
%>
