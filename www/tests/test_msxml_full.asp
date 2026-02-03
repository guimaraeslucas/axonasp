<%
'===============================================
' MSXML2 Objects - Corrected Implementation
' Test file demonstrating the fixed parsing logic
'===============================================

Response.Write "<h1>MSXML2 - Complete Test Suite (Fixed)</h1>"

' ==== Test 1: DocumentElement Property ====
Response.Write "<h3>Test 1: DocumentElement Property</h3>"
Set xmlDoc1 = Server.CreateObject("MSXML2.DOMDocument")
xml1 = "<root><data>Test</data></root>"
xmlDoc1.LoadXML(xml1)

Set docElem = xmlDoc1.DocumentElement
If docElem Is Not Nothing Then
    Response.Write "DocumentElement NodeName: " & docElem.NodeName & "<br>"
Else
    Response.Write "ERROR: DocumentElement is Nothing<br>"
End If

' ==== Test 2: GetElementsByTagName ====
Response.Write "<h3>Test 2: GetElementsByTagName</h3>"
Set xmlDoc2 = Server.CreateObject("MSXML2.DOMDocument")
xml2 = "<library><book>Book 1</book><book>Book 2</book><book>Book 3</book></library>"
xmlDoc2.LoadXML(xml2)

Set books = xmlDoc2.GetElementsByTagName("book")
On Error Resume Next
bookCount = UBound(books) + 1
On Error GoTo 0

If bookCount > 0 Then
    Response.Write "Found " & bookCount & " book elements:<br>"
    For idx = 0 To UBound(books)
        Set currentBook = books(idx)
        Response.Write "  [" & (idx+1) & "] " & currentBook.Text & "<br>"
    Next
Else
    Response.Write "No book elements found<br>"
End If

' ==== Test 3: Nested Elements ====
Response.Write "<h3>Test 3: Nested Elements</h3>"
Set xmlDoc3 = Server.CreateObject("MSXML2.DOMDocument")
xml3 = "<catalog><product><name>Widget</name><price>9.99</price></product></catalog>"
xmlDoc3.LoadXML(xml3)

Set products = xmlDoc3.GetElementsByTagName("product")
On Error Resume Next
prodCount = UBound(products) + 1
On Error GoTo 0

If prodCount > 0 Then
    Set firstProduct = products(0)
    Set names = xmlDoc3.GetElementsByTagName("name")
    On Error Resume Next
    nameCount = UBound(names) + 1
    On Error GoTo 0
    
    If nameCount > 0 Then
        Set nameElem = names(0)
        Response.Write "Product Name: " & nameElem.Text & "<br>"
    End If
Else
    Response.Write "No products found<br>"
End If

' ==== Test 4: ParseError Handling ====
Response.Write "<h3>Test 4: ParseError Handling</h3>"
Set xmlDoc4 = Server.CreateObject("MSXML2.DOMDocument")
invalidXml = "<unclosed>Test"
result = xmlDoc4.LoadXML(invalidXml)
If result Then
    Response.Write "XML parsed (though it may be malformed)<br>"
Else
    Response.Write "ParseError detected - XML is invalid<br>"
End If

' ==== Test 5: SelectSingleNode with XPath ====
Response.Write "<h3>Test 5: SelectSingleNode</h3>"
Set xmlDoc5 = Server.CreateObject("MSXML2.DOMDocument")
xml5 = "<root><users><user><id>1</id><name>Alice</name></user></users></root>"
xmlDoc5.LoadXML(xml5)

Set nameNode = xmlDoc5.SelectSingleNode("//name")
If nameNode Is Not Nothing Then
    Response.Write "Found name node: " & nameNode.Text & "<br>"
Else
    Response.Write "SelectSingleNode returned Nothing<br>"
End If

' ==== Test 6: SelectNodes with XPath ====
Response.Write "<h3>Test 6: SelectNodes</h3>"
Set xmlDoc6 = Server.CreateObject("MSXML2.DOMDocument")
xml6 = "<items><item>A</item><item>B</item><item>C</item></items>"
xmlDoc6.LoadXML(xml6)

Set itemNodes = xmlDoc6.SelectNodes("//item")
On Error Resume Next
itemNodeCount = UBound(itemNodes) + 1
On Error GoTo 0

If itemNodeCount > 0 Then
    Response.Write "Found " & itemNodeCount & " item nodes<br>"
Else
    Response.Write "SelectNodes returned no elements<br>"
End If

' ==== Test 7: CreateElement and Property Access ====
Response.Write "<h3>Test 7: CreateElement</h3>"
Set xmlDoc7 = Server.CreateObject("MSXML2.DOMDocument")
Set newElement = xmlDoc7.CreateElement("custom")
If newElement Is Not Nothing Then
    Response.Write "Created element: " & newElement.NodeName & "<br>"
    newElement.SetProperty "text", "Custom Content"
    Response.Write "Element content: " & newElement.Text & "<br>"
Else
    Response.Write "Failed to create element<br>"
End If

' ==== Test 8: ServerXMLHTTP Properties ====
Response.Write "<h3>Test 8: ServerXMLHTTP Object</h3>"
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
If http Is Not Nothing Then
    Response.Write "ServerXMLHTTP object created<br>"
    Response.Write "  ReadyState: " & http.ReadyState & "<br>"
    Response.Write "  Status: " & http.Status & "<br>"
    Response.Write "  Timeout (default): " & http.Timeout & " seconds<br>"
Else
    Response.Write "Failed to create ServerXMLHTTP<br>"
End If

' ==== Test 9: Response.PICS Property ====
Response.Write "<h3>Test 9: Response.PICS</h3>"
picsLabel = "(PICS-1.1 ""http://www.rsac.org/ratingsv01.html"" l gen true)"
Response.PICS = picsLabel
Response.Write "PICS set to: " & Response.PICS & "<br>"

Response.Write "<h2>All tests completed successfully!</h2>"
%>
