# Axuserhomedirpath Method

## Overview

Returns the current operating system user home directory path.

## Syntax

```asp
result = obj.Axuserhomedirpath()
```

## Return Values

- String: Home directory path, or an empty string when unavailable.

## Code Example

```asp
<%
Option Explicit
Dim obj, p
Set obj = Server.CreateObject("G3AXON.Functions")
p = obj.Axuserhomedirpath()
Response.Write Server.HTMLEncode(p)
Set obj = Nothing
%>
```