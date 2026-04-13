# axfiltervalidateemail

## Overview

The `axfiltervalidateemail` method validates whether a given string follows a valid email address format.

## Syntax

```asp
result = obj.axfiltervalidateemail(emailAddress)
```

## Parameters and Arguments

- **emailAddress** (String): The string to validate as an email address.

## Return Values

Returns a Boolean indicating whether the string is a syntactically valid email address. Returns `True` if valid, otherwise `False`.

## Remarks

- This method is part of the G3Pix AxonASP library.
- Validation is based on standard email syntax rules.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, email
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

email = "user@example.com"

If ax.axfiltervalidateemail(email) Then
    Response.Write "Email is valid."
Else
    Response.Write "Email is invalid."
End If

Set ax = Nothing
%>
```
