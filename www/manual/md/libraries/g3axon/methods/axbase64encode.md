# Encode String to Base64

## Overview

Encodes a string into Base64 format.

## Syntax

```vbscript
strEncoded = obj.axbase64encode(str)
```

## Parameters

- **str** (String): The string to encode.

## Return Value

String. The Base64 encoded representation of the input string.

## Remarks

Base64 encoding is commonly used to safely transmit binary or text data over text-based protocols.

## Code Example

```vbscript
Dim obj, strEncode
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
strEncode = obj.axbase64encode("Hello World")
Response.Write strEncode ' Outputs: SGVsbG8gV29ybGQ=
```