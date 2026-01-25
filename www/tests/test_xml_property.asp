<%
Response.Write "<h1>XML Property Debug</h1>"

Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
xmlStr = "<root><item>Test</item><item2>Data</item2></root>"
result = xmlDoc.LoadXML(xmlStr)

Response.Write "Original XML: " & Server.HTMLEncode(xmlStr) & "<br>"
Response.Write "LoadXML result: " & result & "<br>"
Response.Write "xmlDoc.XML length: " & Len(xmlDoc.XML) & "<br>"
Response.Write "xmlDoc.XML value: [" & Server.HTMLEncode(xmlDoc.XML) & "]<br>"
%>
