# Use AxonASP Library Objects

## Overview
This page lists custom and compatibility objects exposed through Server.CreateObject in AxonASP.

## Syntax
```asp
Set obj = Server.CreateObject("ProgID")
`````

## Parameters and Arguments
- G3MD
- G3CRYPTO and algorithm aliases resolved by runtime
- G3AXON and G3Axon.Functions
- G3JSON
- G3DB
- G3HTTP and G3HTTP.Functions
- G3Mail and compatibility aliases (CDONTS.NewMail, CDO.Message, Persits.MailSender)
- G3Image
- G3FILES
- G3Template
- G3Zip
- G3ZLIB
- G3TAR
- G3ZSTD
- G3FC
- WScript.Shell
- ADOX.Catalog
- MSWC.AdRotator
- MSWC.BrowserType
- MSWC.NextLink
- MSWC.ContentRotator
- MSWC.Counters
- MSWC.PageCounter
- MSWC.Tools
- MSWC.MyInfo
- MSWC.PermissionChecker
- MSXML2.ServerXMLHTTP, MSXML2.XMLHTTP, Microsoft.XMLHTTP
- MSXML2.DOMDocument, Microsoft.XMLDOM
- G3PDF
- G3FileUploader and upload compatibility aliases
- Scripting.FileSystemObject
- Scripting.Dictionary
- ADODB.Stream
- ADODB.Connection
- ADODBOLE.Connection
- ADODB.Recordset
- ADODB.Command
- VBScript.RegExp and RegExp

## Return Values
Server.CreateObject returns a native object value handled by the AxonASP VM dispatch layer.

## Remarks
- Object creation is case-insensitive for supported ProgIDs.
- Unsupported ProgIDs fall through to host-level CreateObject handling.
- Compatibility aliases are mapped to native AxonASP implementations where available.

## Code Example
```asp
<%
Dim db, json, fso
Set db = Server.CreateObject("G3DB")
Set json = Server.CreateObject("G3JSON")
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Response.Write TypeName(db)
%>
`````
