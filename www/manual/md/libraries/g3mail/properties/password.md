# Password Property

## Overview
The **Password** property sets the authentication password for the SMTP server for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.Password
mail.Password = newValue
```

## Parameters and Arguments
- **newValue** (String): The password for the SMTP account.

## Return Values
Returns a **String** containing the current password (if set).

## Remarks
- For security reasons, avoid hardcoding passwords in scripts. Use environment variables (`SMTP_PASS`) or encrypted configuration files.
- This property can also be accessed using the aliases **Pass** or **AuthPassword**.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.Password = "mySecretPassword"

Set mail = Nothing
%>
```
