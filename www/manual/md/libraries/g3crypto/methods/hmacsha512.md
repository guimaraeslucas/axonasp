# HmacSha512 Method

## Overview

Computes an HMAC (Hash-based Message Authentication Code) using the SHA-512 hash function and a secret key with the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
result = obj.HmacSha512(data, key)
```

## Parameters

- **data** (String): The message content to be hashed.
- **key** (String): The secret cryptographic key used for HMAC calculation.

## Return Values

Returns a String containing the calculated HMAC-SHA512 result encoded as a lowercase hexadecimal string.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- SHA-512 provides a stronger security level compared to SHA-256 for systems requiring higher integrity guarantees.
- The resulting hash is also stored in the `Hash` property as a byte array.

## Code Example

```asp
<%
Dim crypto, signature
Set crypto = Server.CreateObject("G3CRYPTO")
signature = crypto.HmacSha512("Hello World", "mySecretKey")
Response.Write "HMAC Signature: " & signature
Set crypto = Nothing
%>
```
