# Level Property

## Overview

The Level property is exposed by the G3ZSTD library object and returns the current state/value associated with this member.

## Syntax

```asp
value = obj.Level
obj.Level = newValue
`````

## Parameters and Arguments

- Getter: no arguments.
- Setter (when supported): one Variant value.

## Return Values

Returns the current property value as Variant. Read-only members reject assignments.

## Remarks

- Property names are case-insensitive.
- Setters are validated by dispatch logic and can raise runtime errors.
- For object-typed values, assign with Set.

## Code Example

```asp
<%
Option Explicit
Dim obj, value
Set obj = Server.CreateObject("G3ZSTD")
value = obj.Level
Response.Write CStr(value)
Set obj = Nothing
%>
`````

