# Cc Property

## Overview
The **Cc** property gets or sets the carbon copy (CC) recipient list for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.Cc
mail.Cc = newValue
```

## Parameters and Arguments
- **newValue** (String): A comma-separated or semicolon-separated list of email addresses.

## Return Values
Returns a **String** containing the current CC recipients, separated by commas.

## Remarks
- Recipients in the CC list are visible to all other recipients.
- Setting this property replaces any existing CC recipients. Use the **AddCc** method to append to the list instead.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.Cc = "cc1@example.com, cc2@example.com"
Response.Write "CC List: " & mail.Cc

Set mail = Nothing
%>
```
