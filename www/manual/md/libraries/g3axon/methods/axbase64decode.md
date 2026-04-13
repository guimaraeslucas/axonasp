# Decode Base64 String

## Overview

Decodes a Base64 encoded string back to its original value.

## Syntax

```vbscript
strDecoded = obj.axbase64decode(str)
```

## Parameters

- **str** (String): The Base64 string to decode.

## Return Value

String. The decoded original string.

## Remarks

If the input string is not a valid Base64 string, an empty string is returned.

## Code Example

```vbscript
Dim obj, strDecode
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
strDecode = obj.axbase64decode("SGVsbG8gV29ybGQ=")
Response.Write strDecode ' Outputs: Hello World
```