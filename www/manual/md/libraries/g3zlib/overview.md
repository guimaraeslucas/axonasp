# Use G3ZLIB in G3Pix AxonASP

## Overview
G3ZLIB is a compression library that allows you to compress and decompress data using the ZLIB format.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("G3ZLIB")
```

## Parameters and Arguments
- ProgID (String, Required): Use `G3ZLIB` to instantiate this object.
- Member access (Optional): Use the methods and properties provided by the library.

## Return Values
Returns a native object handle for the G3ZLIB library.

## Remarks
- Member names are case-insensitive.
- Runtime validation is enforced by object dispatch logic.
- See the methods and properties pages for member-level coverage.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("G3ZLIB")
Response.Write TypeName(obj)
Set obj = Nothing
%>
```
