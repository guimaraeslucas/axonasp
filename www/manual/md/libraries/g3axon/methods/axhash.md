# Compute String Hash by Algorithm

## Overview

Computes a hash of a string using the specified cryptographic algorithm.

## Syntax

```vbscript
strHash = obj.axhash(algo, str)
```

## Parameters

- **algo** (String): The hashing algorithm to use. Supported values are `"md5"`, `"sha1"`, and `"sha256"`.
- **str** (String): The string to hash.

## Return Value

String. The hexadecimal hash of the string.

## Remarks

If an unsupported algorithm is provided, an empty string is returned.

## Code Example

```vbscript
Dim obj, strHash
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
strHash = obj.axhash("sha256", "securepassword")
Response.Write strHash
```