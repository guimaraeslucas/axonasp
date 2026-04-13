# Md5 Method

## Overview

Computes an MD5 (Message Digest 5) hash from the provided input string or byte array using the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
result = obj.Md5(input)
```

## Parameters

- **input** (String or Array): The data to be hashed.

## Return Values

Returns a String containing the 128-bit hash result encoded as a lowercase hexadecimal string.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- MD5 is widely used for checksums and basic data integrity verification.
- **Security Note:** MD5 is no longer considered cryptographically secure for high-value applications or password hashing; use SHA-256 or bcrypt instead.

## Code Example

```asp
<%
Dim crypto, hash
Set crypto = Server.CreateObject("G3CRYPTO")
hash = crypto.Md5("Hello World")
Response.Write "MD5 Hash: " & hash
Set crypto = Nothing
%>
```
