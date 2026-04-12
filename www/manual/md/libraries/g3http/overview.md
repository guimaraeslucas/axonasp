# Use G3HTTP in AxonASP

## Overview
HTTP request helper library.

## Syntax
```asp
Set obj = Server.CreateObject("G3HTTP or G3HTTP.Functions")
`````

## Parameters and Arguments
- ProgID (String, Required): Use one of the supported ProgID forms for this object family.
- Member access (Optional): Use documented method/property members from the library reference pages.

## Return Values
Returns a native object handle for this object family.

## Remarks
- Member names are case-insensitive.
- Runtime validation is enforced by object dispatch logic.
- See the central library methods/properties pages for member-level coverage.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("G3HTTP or G3HTTP.Functions")
Response.Write TypeName(obj)
%>
`````

