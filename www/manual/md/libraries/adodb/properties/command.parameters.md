# Command.Parameters Property

## Overview

The Command.Parameters property is exposed by the ADODB.Connection object in AxonASP.

## Syntax

```asp
value = obj.Command.Parameters
obj.Command.Parameters = newValue
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
value = obj.Command.Parameters
Response.Write CStr(value)
Set obj = Nothing
%>
```