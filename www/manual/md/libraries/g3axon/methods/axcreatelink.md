# Axcreatelink Method

## Overview

Creates a hard link from a source file path to a destination link path.

## Syntax

```asp
result = obj.Axcreatelink(sourcePath, linkPath)
```

## Parameters

- sourcePath (String): Existing source file path.
- linkPath (String): New link path.

## Return Values

- Boolean: True when the operation succeeds; otherwise False.

## Remarks

- Some restricted environments can deny link creation.

## Code Example

```asp
<%
Option Explicit
Dim obj, ok
Set obj = Server.CreateObject("G3AXON.Functions")
ok = obj.Axcreatelink("C:\\temp\\sample.txt", "C:\\temp\\sample.link")
Response.Write CStr(ok)
Set obj = Nothing
%>
```