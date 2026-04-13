# SetBCryptCost Method

## Overview

Configures the computational cost (work factor) for the bcrypt password hashing algorithm using the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
success = obj.SetBCryptCost(cost)
```

## Parameters

- **cost** (Integer): The desired work factor. Accepted values are between 4 and 31.

## Return Values

Returns a Boolean value indicating whether the new cost was successfully applied. Returns `True` if the cost is within the valid range, and `False` otherwise.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- The default cost is 10.
- Increasing the cost exponentially increases the time required to compute a hash, providing greater resistance against brute-force attacks at the expense of server resources.

## Code Example

```asp
<%
Dim crypto, success
Set crypto = Server.CreateObject("G3CRYPTO")
' Increase work factor for higher security
success = crypto.SetBCryptCost(12)
If success Then
    Response.Write "BCrypt cost successfully updated to 12"
End If
Set crypto = Nothing
%>
```
