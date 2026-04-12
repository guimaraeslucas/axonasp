# Use MSXML2 Family in AxonASP

## Overview
AxonASP provides MSXML2 compatibility wrappers for HTTP and XML DOM operations.

## Supported ProgIDs
- MSXML2.ServerXMLHTTP
- MSXML2.DOMDocument
- MSXML2.DOMDocument.3.0
- MSXML2.DOMDocument.6.0
- Microsoft.XMLDOM

## Documentation Map
- Methods index: [Methods](methods.md)
- Properties index: [Properties](properties.md)
- Member details: methods and properties subfolders.

## Code Example
```asp
<%
Dim http, doc
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
http.Open "GET", "https://example.com", False
http.Send

Set doc = Server.CreateObject("MSXML2.DOMDocument")
If doc.LoadXML(http.ResponseText) Then
    If Not doc.DocumentElement Is Nothing Then
        Response.Write doc.DocumentElement.NodeName
    End If
End If
%>
```