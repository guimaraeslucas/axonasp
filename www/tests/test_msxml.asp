<%
Response.Write "<h1>MSXML2 Objects Test</h1>"
Response.Write "<h2>ServerXMLHTTP and DOMDocument</h2>"

' Test 1: ServerXMLHTTP
Response.Write "<h3>Test 1: ServerXMLHTTP</h3>"
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
Response.Write "ServerXMLHTTP object created successfully<br>"

' Test basic properties and methods
http.Open "GET", "http://httpbin.org/get", False
Response.Write "HTTP.Open called<br>"

http.SetRequestHeader "User-Agent", "ASP.NET"
Response.Write "HTTP.SetRequestHeader called<br>"

' Send request
http.Send
Response.Write "HTTP.Send called<br>"

' Check status
Response.Write "HTTP Status: " & http.Status & "<br>"
Response.Write "HTTP StatusText: " & http.StatusText & "<br>"

' Check response
If http.Status = 200 Then
    Response.Write "Response Text (first 100 chars): " & Left(http.ResponseText, 100) & "...<br>"
Else
    Response.Write "HTTP request failed with status: " & http.Status & "<br>"
End If

' Test 2: DOMDocument
Response.Write "<h3>Test 2: DOMDocument - XML Creation and Parsing</h3>"

Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
Response.Write "DOMDocument object created successfully<br>"

' Test LoadXML
xmlString = "<?xml version=""1.0""?><root><item><name>Test Item</name><value>123</value></item></root>"
If xmlDoc.LoadXML(xmlString) Then
    Response.Write "XML loaded successfully<br>"
Else
    Response.Write "Failed to load XML<br>"
End If

' Test GetElementsByTagName
Set itemList = xmlDoc.GetElementsByTagName("item")
Dim itemCount
itemCount = UBound(itemList) + 1
Response.Write "Found " & itemCount & " item elements<br>"

' Test element properties
If itemCount > 0 Then
    Set item = itemList(0)
    Response.Write "Item Name: " & item.NodeName & "<br>"
Else
    Response.Write "No item elements found<br>"
End If

' Test 3: DOMDocument - GetElementsByTagName
Response.Write "<h3>Test 3: GetElementsByTagName</h3>"

Set nameElements = xmlDoc.GetElementsByTagName("name")
Dim nameCount
nameCount = UBound(nameElements) + 1
Response.Write "Found " & nameCount & " name elements<br>"

If nameCount > 0 Then
    Set nameElem = nameElements(0)
    Response.Write "Name Element Value: " & nameElem.Text & "<br>"
Else
    Response.Write "No name elements found<br>"
End If

' Test 4: DOMDocument - CreateElement and appendChild
Response.Write "<h3>Test 4: CreateElement and appendChild</h3>"

Set newElem = xmlDoc.CreateElement("newitem")
Response.Write "Created new element: newitem<br>"

Set nameNode = xmlDoc.CreateElement("description")
Response.Write "Created description element<br>"

If newElem Is Not Nothing Then
    Response.Write "NewElement is not nothing<br>"
Else
    Response.Write "NewElement is nothing<br>"
End If

' Test 5: XML Property Output
Response.Write "<h3>Test 5: Document XML Output</h3>"
Response.Write "XML Content:<br>"
Response.Write "<pre>" & Server.HTMLEncode(xmlDoc.XML) & "</pre>"

' Test 6: ServerXMLHTTP with timeout
Response.Write "<h3>Test 6: ServerXMLHTTP Timeout Property</h3>"
Set http2 = Server.CreateObject("MSXML2.ServerXMLHTTP")
Response.Write "Default timeout: " & http2.Timeout & " seconds<br>"

' Test 7: ParseError property
Response.Write "<h3>Test 7: ParseError Handling</h3>"
Set xmlDoc2 = Server.CreateObject("MSXML2.DOMDocument")
invalidXML = "<root><item>Missing closing tag</root>"
If Not xmlDoc2.LoadXML(invalidXML) Then
    Response.Write "Parse error detected (expected for malformed XML)<br>"
End If

' Test 8: SelectSingleNode (simple XPath)
Response.Write "<h3>Test 8: SelectSingleNode</h3>"
Set singleNode = xmlDoc.SelectSingleNode("//name")
If singleNode Is Not Nothing Then
    Response.Write "Found node: " & singleNode.NodeName & "<br>"
    Response.Write "Node value: " & singleNode.Text & "<br>"
Else
    Response.Write "SelectSingleNode returned nothing<br>"
End If

' Test 9: SelectNodes (simple XPath)
Response.Write "<h3>Test 9: SelectNodes</h3>"
Set nodeList = xmlDoc.SelectNodes("//value")
Dim nodeCount
nodeCount = UBound(nodeList) + 1
Response.Write "Found " & nodeCount & " value nodes<br>"

If nodeCount > 0 Then
    Response.Write "First value node: " & nodeList(0).Text & "<br>"
Else
    Response.Write "No value nodes found<br>"
End If

' Test 10: PICS Property (intrinsic Response object)
Response.Write "<h3>Test 10: Response.PICS Property</h3>"
Response.PICS = "(PICS-1.1 ""http://www.rsac.org/ratingsv01.html"" l gen true comment ""Test PICS Label"")"
Response.Write "PICS Set Successfully<br>"
Response.Write "Current PICS: " & Response.PICS & "<br>"

Response.Write "<h2>All tests completed!</h2>"
%>
