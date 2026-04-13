# AddAddress Method

## Overview
The **AddAddress** method appends a recipient email address to the primary destination list (To) of the G3Pix AxonASP G3MAIL object.

## Syntax
```asp
result = mail.AddAddress(email)
```

## Parameters and Arguments
- **email** (String, Required): The recipient's email address (e.g., "user@example.com").

## Return Values
Returns a **Boolean** value (True).

## Remarks
- You can call this method multiple times to add multiple recipients to the **To** list.
- Invalid email formats are not validated by this method; validation occurs during the **Send** operation.
- The method is case-insensitive and can also be called as **AddRecipient** or **AddTo**.

## Code Example
```asp
<%
Dim mail
Set mail = Server.CreateObject("G3MAIL")

' Add multiple recipients
mail.AddAddress "admin@example.com"
mail.AddAddress "support@example.com"

' Alternatively using the property
mail.To = "user1@example.com, user2@example.com"

Set mail = Nothing
%>
```
