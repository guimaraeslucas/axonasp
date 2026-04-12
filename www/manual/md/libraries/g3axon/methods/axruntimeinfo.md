# Axruntimeinfo Method

## Overview

Returns a phpinfo-style diagnostic report that includes runtime details, server context, memory snapshot, loaded configuration keys from config/axonasp.toml, and the AxonASP legal attribution block.

## Syntax

```asp
result = obj.Axruntimeinfo()
```

## Return Values

- String: Multi-line diagnostic report.

## Remarks

- Method names are case-insensitive.
- This output is intended for diagnostics and administration pages.

## Code Example

```asp
<%
Option Explicit
Dim obj, report
Set obj = Server.CreateObject("G3AXON.Functions")
report = obj.Axruntimeinfo()
Response.Write "<pre>" & Server.HTMLEncode(report) & "</pre>"
Set obj = Nothing
%>
```