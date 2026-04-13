# RandomHex Method

## Overview

Generates cryptographically secure random bytes and returns them as a hexadecimal string using the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
result = obj.RandomHex(size)
```

## Parameters

- **size** (Integer, Optional): The number of random bytes to generate. Defaults to 32.

## Return Values

Returns a String containing the random bytes encoded as a lowercase hexadecimal string.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- A 32-byte hexadecimal string will be 64 characters long because each byte is represented by two hex digits.
- Useful for generating API tokens or session identifiers.

## Code Example

```asp
<%
Dim crypto, randStr
Set crypto = Server.CreateObject("G3CRYPTO")
' Generate 16 bytes encoded as hex (32 characters)
randStr = crypto.RandomHex(16)
Response.Write "Secure Random Hex: " & randStr
Set crypto = Nothing
%>
```
