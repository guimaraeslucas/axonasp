# Port Property

## Overview
The **Port** property gets or sets the port number for the SMTP server for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.Port
mail.Port = newValue
```

## Parameters and Arguments
- **newValue** (Integer): The port number (typically 25, 465, or 587).

## Return Values
Returns an **Integer** representing the current port number.

## Remarks
- If not set, the library attempts to use the `SMTP_PORT` environment variable.
- This property can also be accessed using the alias **SMTPServerPort**.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.Port = 587
Response.Write "SMTP Port: " & mail.Port

Set mail = Nothing
%>
```
