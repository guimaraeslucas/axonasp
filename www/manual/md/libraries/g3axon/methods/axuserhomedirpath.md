# Axuserhomedirpath

## Overview

Returns the current operating system user home directory path.

## Syntax

```asp
result = obj.Axuserhomedirpath()
```

## Return Values

- Returns a String.


## Code Example

```asp
<%
Option Explicit
Dim obj, p
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
p = obj.Axuserhomedirpath()
Response.Write Server.HTMLEncode(p)
Set obj = Nothing
%>
```