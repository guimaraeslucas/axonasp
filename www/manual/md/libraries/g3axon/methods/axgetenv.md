# AxGetEnv

## Overview

Use `AxGetEnv` to read one operating-system environment variable from the host process.

## Prerequisites

- Instantiate the library with `Server.CreateObject("G3AXON.FUNCTIONS")`.

## Syntax

```asp
value = obj.AxGetEnv(name)
```

## Parameters

- **name** (String): Environment variable name to resolve.

## Return Value

- **String**: Returns the value of the requested environment variable.
- **String**: Returns an empty string when the variable does not exist.
- **String**: Returns an empty string when `name` is omitted.

## Remarks

- On Windows hosts, environment variable lookup is case-insensitive.
- On Unix-like hosts, environment variable lookup is case-sensitive.
- Method names are case-insensitive in VBScript dispatch.

## Example

```asp
<%
Option Explicit
Dim ax, pathValue, missingValue

Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

pathValue = ax.AxGetEnv("PATH")
missingValue = ax.AxGetEnv("AXONASP_VAR_THAT_DOES_NOT_EXIST")

Response.Write "PATH length: " & Len(pathValue) & "<br>"
Response.Write "Missing variable value: [" & missingValue & "]"

Set ax = Nothing
%>
```

## API Reference

- **Object**: `G3AXON.FUNCTIONS`
- **Method**: `AxGetEnv`
- **Arguments**: `name As String`
- **Returns**: `String` (environment variable value, or empty string)


