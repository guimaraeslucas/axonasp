# Axchangeowner Method

## Overview

Changes file owner and group identifiers.

## Syntax

```asp
result = obj.Axchangeowner(path, uid, gid)
```

## Parameters

- path (String): Target file path.
- uid (Numeric): User ID.
- gid (Numeric): Group ID.

## Return Values

- Boolean: True when the operation succeeds; otherwise False.

## Remarks

- On Windows or non-privileged contexts, this commonly returns False.

## Code Example

```asp
<%
Option Explicit
Dim obj, ok
Set obj = Server.CreateObject("G3AXON.Functions")
ok = obj.Axchangeowner("/tmp/sample.txt", 0, 0)
Response.Write CStr(ok)
Set obj = Nothing
%>
```