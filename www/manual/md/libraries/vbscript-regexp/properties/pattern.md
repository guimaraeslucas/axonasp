# Pattern Property

## Overview

The Pattern property is exposed by the VBScript.RegExp library object and returns the current state/value associated with this member.

## Syntax

```asp
value = obj.Pattern
obj.Pattern = newValue
```

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
Set obj = Server.CreateObject("VBScript.RegExp")
value = obj.Pattern
Response.Write CStr(value)
Set obj = Nothing
%>
```