<%
Response.Write "<h1>CreateElement Test Only</h1>"

Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")

Response.Write "Test: CreateElement<br>"
Set newElem = xmlDoc.CreateElement("test")
Response.Write "Done!<br>"

If newElem Is Nothing Then
    Response.Write "Result is Nothing<br>"
Else
    Response.Write "NodeName: " & newElem.NodeName & "<br>"
End If
%>
