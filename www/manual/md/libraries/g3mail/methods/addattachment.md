# AddAttachment Method

## Overview
The **AddAttachment** method attaches a file to the email payload of the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
result = mail.AddAttachment(filepath)
```

## Parameters and Arguments
- **filepath** (String, Required): The absolute file path on the server of the file to attach.

## Return Values
Returns a **Boolean** (True on success).

## Remarks
- If the file at **filepath** does not exist, the method raises a File Not Found error.
- Multiple attachments can be added by calling this method multiple times.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

' Attach a document
mail.AddAttachment "C:\inetpub\wwwroot\assets\invoice.pdf"

Set mail = Nothing
%>
```
