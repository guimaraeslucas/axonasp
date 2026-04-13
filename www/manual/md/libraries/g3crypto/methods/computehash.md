# ComputeHash Method

## Overview

Computes a cryptographic hash using the configured or specified algorithm and returns the result as a byte array using the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
result = obj.ComputeHash(input, [algorithm])
```

## Parameters

- **input** (String or Array): The data to be hashed.
- **algorithm** (String, Optional): The hash algorithm to use (e.g., "sha256", "md5"). If omitted, the default algorithm configured during initialization or "sha256" is used.

## Return Values

Returns a VBScript Array of Byte values (VT_ARRAY | VT_UI1) containing the calculated hash.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- The result of the last hash calculation is also stored in the `Hash` property.
- Supported algorithms include md5, sha1, sha256, sha384, sha512, sha3_256, sha3_512, blake2b256, and blake2b512.

## Code Example

```asp
<%
Dim crypto, hashArray
Set crypto = Server.CreateObject("G3CRYPTO")
hashArray = crypto.ComputeHash("Hello World", "sha256")
Response.Write "Hash size: " & UBound(hashArray) + 1 & " bytes"
Set crypto = Nothing
%>
```
