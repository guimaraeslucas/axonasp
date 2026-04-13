# Blake2b256 Method

## Overview

Computes a BLAKE2b-256 hash from the provided input string or byte array using the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
result = obj.Blake2b256(input)
```

## Parameters

- **input** (String or Array): The data to be hashed. Can be a string or a VBScript byte array.

## Return Values

Returns a String containing the 256-bit hash result encoded as a lowercase hexadecimal string.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- BLAKE2b is a cryptographic hash function that is faster than MD5, SHA-1, SHA-2, and SHA-3, yet is at least as secure as the latest standard SHA-3.
- The method name is case-insensitive.

## Code Example

```asp
<%
Dim crypto, hash
Set crypto = Server.CreateObject("G3CRYPTO")
hash = crypto.Blake2b256("Hello AxonASP")
Response.Write "BLAKE2b-256 Hash: " & hash
Set crypto = Nothing
%>
```
