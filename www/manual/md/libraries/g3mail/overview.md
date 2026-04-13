# Use the G3MAIL Library

## Overview
The **G3MAIL** library provides high-performance SMTP (Simple Mail Transfer Protocol) capabilities for G3Pix AxonASP applications. It enables the delivery of electronic mail messages with support for plain text, HTML formatting, multiple recipients (To, CC, BCC), and secure authentication. The library is designed for zero-allocation performance and can leverage system environment variables for default SMTP configuration.

## Syntax
To instantiate the library, use the following syntax:
```asp
Set mail = Server.CreateObject("G3MAIL")
```

## Prerequisites
- **SMTP Server access**: Requires a valid SMTP server hostname or IP address.
- **Authentication**: Valid credentials (Username and Password) are typically required for external relaying.
- **Network access**: The server hosting AxonASP must have outbound access to the SMTP port (usually 25, 465, or 587).

## How it Works
The G3MAIL object operates as a stateful message builder. You configure the server connection details and message properties (such as **Subject**, **Body**, and recipients) before calling the **Send** method. 

If SMTP properties are not explicitly set in the script, the library automatically attempts to retrieve configuration from the following environment variables:
- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_USER`
- `SMTP_PASS`
- `SMTP_FROM`

## API Reference

### Methods
- **AddAddress**: Appends a recipient to the primary destination list.
- **AddBcc**: Appends a recipient to the blind carbon copy list.
- **AddCc**: Appends a recipient to the carbon copy list.
- **Clear**: Resets all message fields and recipient lists to their default state.
- **Send**: Connects to the SMTP server and delivers the message.

### Properties
- **Body**: Sets the message content. Automatically detects format based on **IsHTML**.
- **BodyFormat**: Sets the message format (0 for HTML, 1 for Text).
- **From**: Sets the sender's email address.
- **FromName**: Sets the display name for the sender.
- **Host**: Sets the SMTP server address.
- **IsHTML**: Specifies whether the body should be treated as HTML.
- **Password**: Sets the authentication password.
- **Port**: Sets the SMTP server port.
- **Subject**: Sets the message subject line.
- **To**: Sets the primary recipient list (comma or semicolon separated).
- **Username**: Sets the authentication username.

## Code Example
The following example demonstrates how to configure and send an HTML email.

```asp
<%
Dim mail, result
Set mail = Server.CreateObject("G3MAIL")

' Configure SMTP Server
mail.Host = "smtp.example.com"
mail.Port = 587
mail.Username = "user@example.com"
mail.Password = "securePassword123"

' Configure Message
mail.From = "noreply@example.com"
mail.FromName = "AxonASP Notification"
mail.Subject = "System Alert: Update Completed"
mail.HTMLBody = "<h1>Update Successful</h1><p>Your system has been updated to the latest version.</p>"

' Add Recipients
mail.AddAddress "admin@example.com"
mail.AddCc "backup@example.com"

' Send Email
result = mail.Send()

If result = True Then
    Response.Write "Message delivered successfully."
Else
    Response.Write "Delivery failed: " & result
End If

Set mail = Nothing
%>
```
