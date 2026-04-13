# Bcc Property

## Overview
The **Bcc** property gets or sets the blind carbon copy (BCC) recipient list for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.Bcc
mail.Bcc = newValue
```

## Parameters and Arguments
- **newValue** (String): A comma-separated or semicolon-separated list of email addresses.

## Return Values
Returns a **String** containing the current BCC recipients, separated by commas.

## Remarks
- Recipients in the BCC list are hidden from other recipients.
- Setting this property replaces any existing BCC recipients. Use the **AddBcc** method to append to the list instead.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.Bcc = "audit@example.com, backup@example.com"
Response.Write "BCC List: " & mail.Bcc

Set mail = Nothing
%>
```
