# BodyFormat Property

## Overview
The **BodyFormat** property gets or sets the format of the email message body (Plain Text or HTML) for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.BodyFormat
mail.BodyFormat = newValue
```

## Parameters and Arguments
- **newValue** (Integer): Use 0 for HTML format or 1 for Plain Text format.

## Return Values
Returns an **Integer**: 0 if the message is in HTML format, 1 if it is in Plain Text format.

## Remarks
- This property is synchronized with the **IsHTML** property.
- Setting **BodyFormat** to 0 sets **IsHTML** to True. Setting it to 1 (or any other value) sets **IsHTML** to False.
- This property can also be accessed using the alias **MailFormat**.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.BodyFormat = 0 ' Set to HTML
If mail.BodyFormat = 0 Then
    Response.Write "Message is in HTML format."
End If

Set mail = Nothing
%>
```
