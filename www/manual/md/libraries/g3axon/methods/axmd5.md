# Compute MD5 Hash

## Overview

Computes the MD5 hash of a given string.

## Syntax

```vbscript
strHash = obj.axmd5(str)
```

## Parameters

- **str** (String): The string to hash.

## Return Value

String. The hexadecimal MD5 hash of the string.

## Remarks

MD5 is a widely used cryptographic hash function. It is provided for standard hashing purposes.

## Code Example

```vbscript
Dim obj, strHash
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
strHash = obj.axmd5("admin")
Response.Write strHash ' Outputs: 21232f297a57a5a743894a0e4a801fc3
```