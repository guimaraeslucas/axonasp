# AddCc Method

## Overview
The **AddCc** method appends a recipient email address to the carbon copy (CC) list of the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
result = mail.AddCc(email)
```

## Parameters and Arguments
- **email** (String, Required): The recipient's email address to be CC'd.

## Return Values
Returns a **Boolean** value (True).

## Remarks
- Recipients added via this method will be visible to all other recipients of the email.
- You can call this method multiple times to add multiple CC recipients.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.Subject = "Project Update"
mail.AddAddress "manager@example.com"
mail.AddCc "team-lead@example.com"
mail.AddCc "backup@example.com"

Set mail = Nothing
%>
```
