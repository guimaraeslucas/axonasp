# Username Property

## Overview
The **Username** property gets or sets the authentication username for the SMTP server for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.Username
mail.Username = newValue
```

## Parameters and Arguments
- **newValue** (String): The username or email address for SMTP authentication.

## Return Values
Returns a **String** containing the current username.

## Remarks
- If not set, the library attempts to use the `SMTP_USER` environment variable.
- This property can also be accessed using the aliases **User** or **AuthUsername**.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.Username = "webmaster@example.com"
Response.Write "Auth User: " & mail.Username

Set mail = Nothing
%>
```
