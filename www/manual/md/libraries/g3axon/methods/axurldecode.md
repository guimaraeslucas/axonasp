# Decode URL String

## Overview

Decodes a URL-encoded string, standardizing query string formats.

## Syntax

```vbscript
strDecoded = obj.axurldecode(str)
```

## Parameters

- **str** (String): The URL-encoded string to decode.

## Return Value

String. The decoded string.

## Remarks

Conversely to raw URL decoding, this function handles standard query string encoded inputs properly.

## Code Example

```vbscript
Dim obj, strDecode
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
strDecode = obj.axurldecode("Hello%20World")
Response.Write strDecode ' Outputs: Hello World
```