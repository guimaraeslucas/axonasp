# HtmlBody Property

## Overview
The **HtmlBody** property gets or sets the HTML content of the email message for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.HtmlBody
mail.HtmlBody = newValue
```

## Parameters and Arguments
- **newValue** (String): The HTML content of the email.

## Return Values
Returns a **String** containing the current HTML body.

## Remarks
- Setting this property automatically toggles **IsHTML** to True.
- This property allows the JScript runtime to read it, modify it, and call native string methods (such as `.match()` and `.replace()`) directly on it.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.HtmlBody = "<h1>Hello World</h1><p>This is an HTML body.</p>"

Set mail = Nothing
%>
```
