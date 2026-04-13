# IsHTML Property

## Overview
The **IsHTML** property gets or sets a value indicating whether the email body should be treated as HTML for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.IsHTML
mail.IsHTML = newValue
```

## Parameters and Arguments
- **newValue** (Boolean): True for HTML content, False for plain text.

## Return Values
Returns a **Boolean** value.

## Remarks
- Setting this property directly affects how the **Body** content is rendered by the recipient's mail client.
- This property is synchronized with **BodyFormat**.
- Some aliases like **HtmlBody** set this property to True automatically.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.Body = "<h1>Notice</h1><p>System maintenance tonight.</p>"
mail.IsHTML = True

Set mail = Nothing
%>
```
