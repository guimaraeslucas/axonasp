# Axchangetimes Method

## Overview

Changes file access and modification timestamps using Unix epoch seconds.

## Syntax

```asp
result = obj.Axchangetimes(path, accessTime, modifiedTime)
```

## Parameters

- path (String): Target file path.
- accessTime (Numeric): Access timestamp in Unix epoch seconds.
- modifiedTime (Numeric): Modified timestamp in Unix epoch seconds.

## Return Values

- Boolean: True when the operation succeeds; otherwise False.

## Code Example

```asp
<%
Option Explicit
Dim obj, ok
Set obj = Server.CreateObject("G3AXON.Functions")
ok = obj.Axchangetimes("C:\\temp\\sample.txt", 1700000000, 1700000001)
Response.Write CStr(ok)
Set obj = Nothing
%>
```