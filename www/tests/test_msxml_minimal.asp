<%
Response.Write "<h1>Simple MSXML Test</h1>"

Response.Write "<h3>Step 1: Create Object</h3>"
Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
Response.Write "Created<br>"

Response.Write "<h3>Step 2: Load XML</h3>"
result = xmlDoc.LoadXML("<root><item>Test</item></root>")
Response.Write "LoadXML result: " & result & "<br>"

Response.Write "<h3>Step 3: Get XML property</h3>"
Response.Write "XML: " & xmlDoc.XML & "<br>"

Response.Write "<h3>Done!</h3>"
%>
