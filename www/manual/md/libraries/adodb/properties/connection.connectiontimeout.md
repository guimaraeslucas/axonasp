# Connection.ConnectionTimeout Property

## Overview

The Connection.ConnectionTimeout property is exposed by the ADODB.Connection object in AxonASP.

## Syntax

```asp
value = obj.Connection.ConnectionTimeout
obj.Connection.ConnectionTimeout = newValue
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
value = obj.Connection.ConnectionTimeout
Response.Write CStr(value)
Set obj = Nothing
%>
```