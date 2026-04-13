# Compute SHA-1 Hash

## Overview

Computes the SHA-1 hash of a given string.

## Syntax

```vbscript
strHash = obj.axsha1(str)
```

## Parameters

- **str** (String): The string to hash.

## Return Value

String. The hexadecimal SHA-1 hash of the string.

## Remarks

SHA-1 produces a 160-bit hash value, commonly used for checksums.

## Code Example

```vbscript
Dim obj, strHash
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
strHash = obj.axsha1("test")
Response.Write strHash ' Outputs: a94a8fe5ccb19ba61c4c0873d391e987982fbbd3
```