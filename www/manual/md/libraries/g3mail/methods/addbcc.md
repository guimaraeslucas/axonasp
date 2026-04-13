# AddBcc Method

## Overview
The **AddBcc** method appends a recipient email address to the blind carbon copy (BCC) list of the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
result = mail.AddBcc(email)
```

## Parameters and Arguments
- **email** (String, Required): The recipient's email address to be BCC'd.

## Return Values
Returns a **Boolean** value (True).

## Remarks
- Recipients added via this method are not visible to other recipients of the email.
- Use this method for privacy when sending to multiple recipients who should not see each other's addresses.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.Subject = "Newsletter"
mail.AddAddress "subscriber@example.com"
mail.AddBcc "archive@example.com"
mail.AddBcc "compliance@example.com"

Set mail = Nothing
%>
```
