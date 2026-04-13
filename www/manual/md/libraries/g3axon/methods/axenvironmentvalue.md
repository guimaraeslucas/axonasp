# axenvironmentvalue

## Overview
Retrieves the value of a specific environment variable for the G3Pix AxonASP process.

## Syntax
```asp
result = obj.axenvironmentvalue(name [, default])
```

## Parameters and Arguments
- **name** (String): The name of the environment variable to retrieve.
- **default** (Variant, Optional): The value to return if the environment variable is not found.

## Return Values
Returns a String containing the value of the environment variable. If the variable is not found, it returns the provided default value or an empty string if no default was specified.

## Remarks
The lookup is case-sensitive or case-insensitive depending on the underlying operating system (e.g., case-insensitive on Windows, case-sensitive on Linux).

## Code Example
```asp
<%
Option Explicit
Dim obj, path, nonExistent
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

' Retrieve a common environment variable
path = obj.axenvironmentvalue("PATH")
Response.Write "System PATH: " & path & "<br>"

' Use a default value for a non-existent variable
nonExistent = obj.axenvironmentvalue("MY_CUSTOM_VAR", "DEFAULT_VALUE")
Response.Write "Custom Var: " & nonExistent

Set obj = Nothing
%>
```
