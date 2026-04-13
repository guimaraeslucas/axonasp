# From Property

## Overview
The **From** property gets or sets the sender's email address for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.From
mail.From = newValue
```

## Parameters and Arguments
- **newValue** (String): The email address of the sender.

## Return Values
Returns a **String** containing the sender's email address.

## Remarks
- This property is required for sending emails unless the `SMTP_FROM` environment variable is set.
- To set a display name (e.g., "Company Name <sender@example.com>"), use the **FromName** property.
- This property can also be accessed using the alias **FromAddress**.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.From = "no-reply@example.com"
Response.Write "Sender: " & mail.From

Set mail = Nothing
%>
```
