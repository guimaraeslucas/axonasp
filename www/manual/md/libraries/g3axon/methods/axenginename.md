# AxEngineName

## Overview

Use `AxEngineName` to identify the runtime engine name exposed by G3Pix AxonASP.

## Prerequisites

- Instantiate the library with `Server.CreateObject("G3AXON.FUNCTIONS")`.

## Syntax

```asp
engineName = obj.AxEngineName()
```

## Parameters

- This method does not require parameters.

## Return Value

- **String**: Always returns `AxonASP`.

## Remarks

- Use this method for runtime identification checks before calling engine-specific routines.
- Method names are case-insensitive in VBScript dispatch.

## Example

```asp
<%
Option Explicit
Dim ax, engineName

Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

engineName = ax.AxEngineName()
If engineName = "AxonASP" Then
    Response.Write "Running on G3Pix AxonASP"
Else
    Response.Write "Unknown engine"
End If

Set ax = Nothing
%>
```

## API Reference

- **Object**: `G3AXON.FUNCTIONS`
- **Method**: `AxEngineName`
- **Arguments**: none
- **Returns**: `String` (`AxonASP`)


