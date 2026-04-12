# Fields.Count Property

## Overview

The Fields.Count property is exposed by the ADODB.Connection object in AxonASP.

## Syntax

```asp
value = obj.Fields.Count
obj.Fields.Count = newValue
```
## Parameters and Arguments

- Getter: No arguments.
- Setter (when supported): One Variant value.

## Return Values

Returns the current property value as Variant. Read-only members reject assignments.

## Remarks

- Property names are case-insensitive.
- Setters are validated by runtime dispatch and can raise runtime errors.
- For object-typed values, assign with Set.

## Code Example

```asp
<%
Option Explicit
Dim obj, value
Set obj = Server.CreateObject("ADODB.Connection")
value = obj.Fields.Count
Response.Write CStr(value)
Set obj = Nothing
%>
```