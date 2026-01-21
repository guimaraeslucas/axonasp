<%
Response.Write "<h1>DocumentElement Nothing Test</h1>"

Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
xmlDoc.LoadXML("<root><data>Test</data></root>")

Set docElem = xmlDoc.DocumentElement

Response.Write "Type of docElem: " & TypeName(docElem) & "<br>"
Response.Write "Is Nothing test: " & (docElem Is Nothing) & "<br>"
Response.Write "Not Is Nothing test: " & (Not docElem Is Nothing) & "<br>"

If docElem Is Nothing Then
    Response.Write "docElem Is Nothing = TRUE<br>"
Else
    Response.Write "docElem Is Nothing = FALSE<br>"
End If

If Not docElem Is Nothing Then
    Response.Write "Not docElem Is Nothing = TRUE<br>"
    Response.Write "NodeName: " & docElem.NodeName & "<br>"
Else
    Response.Write "Not docElem Is Nothing = FALSE<br>"
End If
%>
