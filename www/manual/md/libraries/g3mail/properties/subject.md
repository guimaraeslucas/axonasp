# Subject Property

## Overview
The **Subject** property gets or sets the subject line of the email for the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
value = mail.Subject
mail.Subject = newValue
```

## Parameters and Arguments
- **newValue** (String): The text for the email subject.

## Return Values
Returns a **String** containing the current subject.

## Remarks
- The subject line should be concise and relevant to the message content.
- It is highly recommended to set a subject for every email to avoid being flagged as spam.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

mail.Subject = "Urgent: System Report"
Response.Write "Subject: " & mail.Subject

Set mail = Nothing
%>
```
