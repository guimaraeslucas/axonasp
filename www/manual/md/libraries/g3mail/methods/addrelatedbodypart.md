# AddRelatedBodyPart Method

## Overview
The **AddRelatedBodyPart** method embeds a related resource (such as an inline image) into the email payload, mapping it to a specific Content-ID (CID).

## Syntax
```asp
Set bodyPart = mail.AddRelatedBodyPart(filepath, cid)
```

## Parameters and Arguments
- **filepath** (String, Required): The absolute file path of the resource on the server.
- **cid** (String, Required): The Content-ID to assign to the resource.

## Return Values
Returns a body part object that supports the **Fields** collection.

## Remarks
- If the file at **filepath** does not exist, the method raises a File Not Found error.
- The returned body part object exposes a **Fields** property, allowing you to update its headers (such as `urn:schemas:mailheader:Content-ID`) prior to sending.
- In JScript, the Content-ID can be updated via the returned object's fields using the `Fields.Item("urn:schemas:mailheader:Content-ID") = "<" + cid + ">"` chain.

## Code Example
```asp
<%
Dim mail, bp
Set mail = Server.CreateObject("G3MAIL")

' Add related body part
Set bp = mail.AddRelatedBodyPart("C:\inetpub\wwwroot\images\logo.png", "logo")

' In VBScript, the Fields collection can be updated
bp.Fields.Item("urn:schemas:mailheader:Content-ID") = "<logo>"
bp.Fields.Update()

Set mail = Nothing
%>
```
