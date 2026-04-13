# Send Method

## Overview
The **Send** method attempts to deliver the email message using the configured SMTP server and message properties for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
result = mail.Send([to, subject, body])
```

## Parameters and Arguments
- **to** (String, Optional): The recipient email address. If provided, it overrides the **To** property for this call.
- **subject** (String, Optional): The message subject. If provided, it overrides the **Subject** property for this call.
- **body** (String, Optional): The message body. If provided, it overrides the **Body** property for this call and sets the format to plain text.

## Return Values
Returns a **Boolean** value (True) if the email was sent successfully. If the operation fails, it returns a **String** containing the error message.

## Remarks
- If SMTP server properties (Host, Port, etc.) are not set, the library will attempt to use environment variables (`SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`, `SMTP_FROM`).
- At least one recipient must be defined (via **AddAddress**, **AddCc**, **AddBcc**, or the **To** property).
- The method uses the `gomail` library internally for reliable delivery.

## Code Example
```asp
<%
Dim mail, result
Set mail = Server.CreateObject("G3MAIL")

mail.Host = "smtp.example.com"
mail.From = "system@example.com"
mail.Subject = "Test Email"
mail.Body = "This is a test message."
mail.AddAddress "recipient@example.com"

result = mail.Send()

If result = True Then
    Response.Write "Success!"
Else
    Response.Write "Error: " & result
End If

Set mail = Nothing
%>
```
