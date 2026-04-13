# RandomBase64 Method

## Overview

Generates cryptographically secure random bytes and returns them as a Base64-encoded string using the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
result = obj.RandomBase64(size)
```

## Parameters

- **size** (Integer, Optional): The number of random bytes to generate. Defaults to 32.

## Return Values

Returns a String containing the random bytes encoded in Base64 format.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- This method uses a secure random number generator provided by the underlying operating system.
- Base64 encoding results in a string that is approximately 33% longer than the raw byte length.

## Code Example

```asp
<%
Dim crypto, randStr
Set crypto = Server.CreateObject("G3CRYPTO")
' Generate 16 bytes encoded as Base64
randStr = crypto.RandomBase64(16)
Response.Write "Secure Random Base64: " & randStr
Set crypto = Nothing
%>
```
