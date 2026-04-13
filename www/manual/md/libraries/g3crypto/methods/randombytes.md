# RandomBytes Method

## Overview

Generates cryptographically secure random bytes and returns them as a VBScript byte array using the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
result = obj.RandomBytes(size)
```

## Parameters

- **size** (Integer, Optional): The number of random bytes to generate. Defaults to 32.

## Return Values

Returns a VBScript Array of Byte values (VT_ARRAY | VT_UI1) containing the generated random data.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- This method is suitable for generating salts, nonces, and secret keys.
- If an error occurs during random data generation, an empty array is returned.

## Code Example

```asp
<%
Dim crypto, randArray
Set crypto = Server.CreateObject("G3CRYPTO")
' Generate 16 secure random bytes
randArray = crypto.RandomBytes(16)
Response.Write "First random byte: " & randArray(0)
Set crypto = Nothing
%>
```
