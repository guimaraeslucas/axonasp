# Axcachedirpath Method

## Overview

Returns the full cache directory path for .temp/cache/ with a trailing path separator.

## Syntax

```asp
result = obj.Axcachedirpath()
```

## Return Values

- String: Absolute cache directory path ending with path separator.

## Code Example

```asp
<%
Option Explicit
Dim obj, p
Set obj = Server.CreateObject("G3AXON.Functions")
p = obj.Axcachedirpath()
Response.Write Server.HTMLEncode(p)
Set obj = Nothing
%>
```