# Clear Method

## Overview
The **Clear** method resets all message properties and recipient lists of the G3Pix AxonASP G3MAIL object to their default empty states.

## Syntax
```asp
result = mail.Clear()
```

## Parameters and Arguments
This method does not take any arguments.

## Return Values
Returns a **Boolean** value (True).

## Remarks
- This method clears the **To**, **Cc**, and **Bcc** lists, as well as the **Subject**, **Body**, and **IsHTML** status.
- Use this method when reusing the same G3MAIL object to send multiple different emails in a single script execution.
- SMTP server settings (Host, Port, Username, Password) are NOT cleared by this method.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

' First email
mail.AddAddress "user1@example.com"
mail.Subject = "First Message"
mail.Body = "Hello User 1"
mail.Send()

' Clear for second email
mail.Clear()

' Second email
mail.AddAddress "user2@example.com"
mail.Subject = "Second Message"
mail.Body = "Hello User 2"
mail.Send()

Set mail = Nothing
%>
```
