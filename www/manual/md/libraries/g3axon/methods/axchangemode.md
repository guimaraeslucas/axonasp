# Axchangemode Method

## Overview

Changes file mode permissions from an octal text value.

## Syntax

```asp
result = obj.Axchangemode(path, mode)
```

## Parameters

- path (String): Target file path.
- mode (String): Octal mode text, such as 0644.

## Return Values

- Boolean: True when the operation succeeds; otherwise False.

## Code Example

```asp
<%
Option Explicit
Dim obj, ok
Set obj = Server.CreateObject("G3AXON.Functions")
ok = obj.Axchangemode("C:\\temp\\sample.txt", "0644")
Response.Write CStr(ok)
Set obj = Nothing
%>
```