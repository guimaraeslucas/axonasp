# Sha1 Method

## Overview

Computes a SHA-1 (Secure Hash Algorithm 1) hash from the provided input string or byte array using the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
result = obj.Sha1(input)
```

## Parameters

- **input** (String or Array): The data to be hashed.

## Return Values

Returns a String containing the 160-bit hash result encoded as a lowercase hexadecimal string.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- SHA-1 was widely used for integrity checks and digital signatures.
- **Security Note:** SHA-1 is no longer considered secure against well-funded attackers; it is recommended to use SHA-256 or higher for security-sensitive applications.

## Code Example

```asp
<%
Dim crypto, hash
Set crypto = Server.CreateObject("G3CRYPTO")
hash = crypto.Sha1("Hello World")
Response.Write "SHA-1 Hash: " & hash
Set crypto = Nothing
%>
```
