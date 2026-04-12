# Use MSXML2 Compatibility Objects in AxonASP

## Overview
AxonASP supports HTTP and XML DOM compatibility objects modeled after MSXML2 ServerXMLHTTP and DOMDocument behavior.

## Syntax
```asp
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
Set doc = Server.CreateObject("MSXML2.DOMDocument")
`````

## Parameters and Arguments
- ServerXMLHTTP methods: open, setrequestheader, send, abort, getresponseheader, getallresponseheaders.
- ServerXMLHTTP properties: responsetext, responsexml, responsebody, status, statustext, readystate, timeout.
- DOMDocument methods: loadxml, load, save, getelementsbytagname, createelement, createtextnode, createattribute, appendchild, selectsinglenode, selectnodes, getproperty, setproperty.
- DOMDocument properties: documentelement, xml, parseerror, async, serverhttprequest, resolveexternals, validateonparse, preservewhitespace, selectionlanguage, selectionnamespaces.
- XMLNodeList methods: item, nextnode.
- XMLNodeList properties: length, count.
- ParseError properties: errorcode, reason, filepos, line, linepos, srctext, url.
- XMLElement methods: appendchild, getelementsbytagname, item, setattribute, getattribute, removeattribute, selectsinglenode, selectnodes.
- XMLElement properties: nodename, nodevalue, text, xml, attributes, childnodes, firstchild, lastchild, parentnode, length, children.

## Return Values
MSXML object methods return values compatible with AxonASP DOM/HTTP wrappers, including object handles for node collections and elements.

## Remarks
- Object names and members are case-insensitive.
- DOM selection behavior depends on current document state and parser output.
- ParseError values are populated on XML parse/load failures.

## Code Example
```asp
<%
Dim http, doc
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
http.open "GET", "https://example.com/feed.xml", False
http.send

Set doc = Server.CreateObject("MSXML2.DOMDocument")
doc.loadxml http.responseText
If Not doc.documentElement Is Nothing Then
  Response.Write doc.documentElement.nodeName
End If
%>
`````
