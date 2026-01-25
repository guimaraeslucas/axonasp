<%
Response.Write "<h1>Not Operator Precedence Test</h1>"

Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
xmlDoc.LoadXML("<root><data>Test</data></root>")
Set docElem = xmlDoc.DocumentElement

Response.Write "<h3>Test 1: With Parentheses</h3>"
If Not (docElem Is Nothing) Then
    Response.Write "SUCCESS: DocumentElement exists<br>"
    Response.Write "NodeName: " & docElem.NodeName & "<br>"
Else
    Response.Write "FAIL: DocumentElement is Nothing<br>"
End If

Response.Write "<h3>Test 2: Without Parentheses</h3>"
If Not docElem Is Nothing Then
    Response.Write "SUCCESS: DocumentElement exists<br>"
    Response.Write "NodeName: " & docElem.NodeName & "<br>"
Else
    Response.Write "FAIL: DocumentElement is Nothing<br>"
End If

Response.Write "<h3>Test 3: Alternative Syntax</h3>"
If docElem Is Nothing Then
    Response.Write "FAIL: DocumentElement is Nothing<br>"
Else
    Response.Write "SUCCESS: DocumentElement exists<br>"
    Response.Write "NodeName: " & docElem.NodeName & "<br>"
End If
%>
