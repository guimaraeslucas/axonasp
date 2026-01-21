<%
Response.Write "<h1>Test 1 Replica</h1>"

Set xmlDoc1 = Server.CreateObject("MSXML2.DOMDocument")
Response.Write "Created xmlDoc1<br>"

xml1 = "<root><data>Test</data></root>"
result = xmlDoc1.LoadXML(xml1)
Response.Write "LoadXML result: " & result & "<br>"

Response.Write "Getting DocumentElement...<br>"
Set docElem = xmlDoc1.DocumentElement
Response.Write "Got DocumentElement<br>"

If Not docElem Is Nothing Then
    Response.Write "DocumentElement NodeName: " & docElem.NodeName & "<br>"
Else
    Response.Write "ERROR: DocumentElement is Nothing<br>"
End If
%>
