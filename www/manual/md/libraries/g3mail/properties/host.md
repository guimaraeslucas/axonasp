# Host Property

## Overview
The **Host** property gets or sets the SMTP server hostname or IP address for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.Host
mail.Host = newValue
```

## Parameters and Arguments
- **newValue** (String): The address of the SMTP server (e.g., "smtp.gmail.com").

## Return Values
Returns a **String** containing the current SMTP host.

## Remarks
- If not set, the library attempts to use the `SMTP_HOST` environment variable.
- This property can also be accessed using the aliases **MailHost** or **SMTPServer**.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.Host = "smtp.office365.com"
Response.Write "Current Host: " & mail.Host

Set mail = Nothing
%>
```
