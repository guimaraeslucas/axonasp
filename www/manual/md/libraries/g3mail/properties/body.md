# Body Property

## Overview
The **Body** property gets or sets the main content of the email message for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.Body
mail.Body = newValue
```

## Parameters and Arguments
- **newValue** (String): The text or HTML content of the email.

## Return Values
Returns a **String** containing the current message body.

## Remarks
- Setting this property via **Body**, **Message**, or **TextBody** automatically sets **IsHTML** to False.
- Setting this property via **HtmlBody** automatically sets **IsHTML** to True.
- This property is essential for the **Send** method to have content to deliver.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.Body = "Hello, this is a plain text message."
' To send HTML:
' mail.HtmlBody = "<h1>Hello</h1><p>This is HTML.</p>"

Set mail = Nothing
%>
```
