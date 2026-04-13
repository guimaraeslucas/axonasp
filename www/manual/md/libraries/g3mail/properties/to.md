# To Property

## Overview
The **To** property gets or sets the primary recipient list for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.To
mail.To = newValue
```

## Parameters and Arguments
- **newValue** (String): A comma-separated or semicolon-separated list of email addresses.

## Return Values
Returns a **String** containing the current primary recipients, separated by commas.

## Remarks
- At least one recipient must be specified (either here or via **AddAddress**) before calling **Send**.
- Setting this property replaces any existing primary recipients. Use the **AddAddress** method to append to the list instead.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.To = "customer1@example.com, customer2@example.com"
Response.Write "Recipients: " & mail.To

Set mail = Nothing
%>
```
