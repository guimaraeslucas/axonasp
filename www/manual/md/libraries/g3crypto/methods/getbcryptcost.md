# GetBCryptCost Method

## Overview

Retrieves the current bcrypt computational cost used for password hashing using the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
cost = obj.GetBCryptCost()
```

## Parameters

This method accepts no parameters.

## Return Values

Returns an Integer representing the current work factor (cost) used by the bcrypt algorithm.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- The default cost is 10. Higher cost increases the time needed to compute a hash, enhancing security against brute-force attacks.
- This value can also be accessed and modified via the `BCryptCost` property.

## Code Example

```asp
<%
Dim crypto, cost
Set crypto = Server.CreateObject("G3CRYPTO")
cost = crypto.GetBCryptCost()
Response.Write "Current BCrypt cost: " & cost
Set crypto = Nothing
%>
```
