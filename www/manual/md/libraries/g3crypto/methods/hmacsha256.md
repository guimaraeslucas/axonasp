# HmacSha256 Method

## Overview

Computes an HMAC (Hash-based Message Authentication Code) using the SHA-256 hash function and a secret key with the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
result = obj.HmacSha256(data, key)
```

## Parameters

- **data** (String): The message content to be hashed.
- **key** (String): The secret cryptographic key used for HMAC calculation.

## Return Values

Returns a String containing the calculated HMAC-SHA256 result encoded as a lowercase hexadecimal string.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- HMAC provides data integrity and authentication.
- The resulting hash is also stored in the `Hash` property as a byte array.

## Code Example

```asp
<%
Dim crypto, signature
Set crypto = Server.CreateObject("G3CRYPTO")
signature = crypto.HmacSha256("Hello World", "mySecretKey")
Response.Write "HMAC Signature: " & signature
Set crypto = Nothing
%>
```
