# Blake2b512 Method

## Overview

Computes a BLAKE2b-512 hash from the provided input string or byte array using the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
result = obj.Blake2b512(input)
```

## Parameters

- **input** (String or Array): The data to be hashed. Can be a string or a VBScript byte array.

## Return Values

Returns a String containing the 512-bit hash result encoded as a lowercase hexadecimal string.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- BLAKE2b is optimized for 64-bit platforms and yields very high performance on modern CPUs.
- The method name is case-insensitive.

## Code Example

```asp
<%
Dim crypto, hash
Set crypto = Server.CreateObject("G3CRYPTO")
hash = crypto.Blake2b512("Hello AxonASP")
Response.Write "BLAKE2b-512 Hash: " & hash
Set crypto = Nothing
%>
```
