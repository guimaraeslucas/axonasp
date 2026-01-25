<%
Response.Write "<h1>SelectSingleNode Debug</h1>"

Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
xmlDoc.LoadXML("<root><users><user><name>Alice</name></user></users></root>")

Response.Write "XML loaded<br>"

Response.Write "<h3>Test 1: Simple XPath</h3>"
Set node1 = xmlDoc.SelectSingleNode("root")
If node1 Is Nothing Then
    Response.Write "XPath 'root' returned Nothing<br>"
Else
    Response.Write "Found: " & node1.NodeName & "<br>"
End If

Response.Write "<h3>Test 2: Nested XPath</h3>"
Set node2 = xmlDoc.SelectSingleNode("root/users")
If node2 Is Nothing Then
    Response.Write "XPath 'root/users' returned Nothing<br>"
Else
    Response.Write "Found: " & node2.NodeName & "<br>"
End If

Response.Write "<h3>Test 3: Deep XPath</h3>"
Set node3 = xmlDoc.SelectSingleNode("root/users/user/name")
If node3 Is Nothing Then
    Response.Write "XPath 'root/users/user/name' returned Nothing<br>"
Else
    Response.Write "Found: " & node3.NodeName & " = " & node3.Text & "<br>"
End If

Response.Write "<h3>Test 4: XPath with //</h3>"
Set node4 = xmlDoc.SelectSingleNode("//name")
If node4 Is Nothing Then
    Response.Write "XPath '//name' returned Nothing<br>"
Else
    Response.Write "Found: " & node4.NodeName & " = " & node4.Text & "<br>"
End If
%>
