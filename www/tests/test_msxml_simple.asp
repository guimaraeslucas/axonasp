<%
Response.Write "<h1>MSXML2 Objects - Fixed Test</h1>"

' Test 1: Simple XML Creation and Element Finding
Response.Write "<h3>Test 1: XML Parsing and Element Search</h3>"

Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
Response.Write "DOMDocument created<br>"

' Simple XML for testing
simpleXML = "<root><item>First</item><item>Second</item></root>"
If xmlDoc.LoadXML(simpleXML) Then
    Response.Write "XML loaded successfully<br>"
Else
    Response.Write "XML load failed<br>"
End If

' Try to find items
Set items = xmlDoc.GetElementsByTagName("item")
If items Is Not Nothing Then
    ' For array checking, create a test variable
    On Error Resume Next
    testCount = UBound(items)
    On Error GoTo 0
    
    If testCount >= 0 Then
        itemLen = testCount + 1
    Else
        itemLen = 0
    End If
    
    Response.Write "Found " & itemLen & " item(s)<br>"
    
    If itemLen > 0 Then
        For i = 0 To testCount
            Set currItem = items(i)
            If currItem Is Not Nothing Then
                Response.Write "  Item " & (i+1) & ": " & currItem.Text & "<br>"
            End If
        Next
    End If
Else
    Response.Write "GetElementsByTagName returned Nothing<br>"
End If

' Test 2: Root element access
Response.Write "<h3>Test 2: Root Element Access</h3>"
Set root = xmlDoc.DocumentElement
If root Is Not Nothing Then
    Response.Write "Root element name: " & root.NodeName & "<br>"
Else
    Response.Write "DocumentElement is Nothing<br>"
End If

' Test 3: Create and manipulate elements
Response.Write "<h3>Test 3: CreateElement and AppendChild</h3>"
Set newElem = xmlDoc.CreateElement("newitem")
If newElem Is Not Nothing Then
    Response.Write "Created new element: " & newElem.NodeName & "<br>"
    
    Set textNode = xmlDoc.CreateTextNode("Test Content")
    If textNode Is Not Nothing Then
        Response.Write "Created text node<br>"
    End If
Else
    Response.Write "Failed to create element<br>"
End If

' Test 4: PICS property
Response.Write "<h3>Test 4: Response.PICS Property</h3>"
Response.PICS = "(PICS-1.1 ""http://www.rsac.org/ratingsv01.html"")"
Response.Write "PICS set successfully<br>"
Response.Write "PICS value: " & Response.PICS & "<br>"

' Test 5: ServerXMLHTTP basic test
Response.Write "<h3>Test 5: ServerXMLHTTP Object</h3>"
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
If http Is Not Nothing Then
    Response.Write "ServerXMLHTTP created successfully<br>"
    Response.Write "ReadyState: " & http.ReadyState & "<br>"
    Response.Write "Timeout: " & http.Timeout & " seconds<br>"
Else
    Response.Write "Failed to create ServerXMLHTTP<br>"
End If

Response.Write "<h2>Tests completed!</h2>"
%>
