# FromName Property

## Overview
The **FromName** property gets or sets the display name of the sender for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.FromName
mail.FromName = newValue
```

## Parameters and Arguments
- **newValue** (String): The friendly display name of the sender.

## Return Values
Returns a **String** containing the sender's display name.

## Remarks
- When set, the recipient will see this name in their mail client instead of just the email address.
- This property works in conjunction with the **From** property.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.From = "info@g3pix.com.br"
mail.FromName = "G3Pix Support"

Set mail = Nothing
%>
```
