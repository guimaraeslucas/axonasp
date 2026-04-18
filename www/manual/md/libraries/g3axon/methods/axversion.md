# AxVersion

## Overview

Use `AxVersion` to read the active runtime version string of G3Pix AxonASP.

## Prerequisites

- Instantiate the library with `Server.CreateObject("G3AXON.FUNCTIONS")`.

## Syntax

```asp
versionText = obj.AxVersion()
```

## Parameters

- This method does not require parameters.

## Return Value

- **String**: Returns the current AxonASP runtime version (for example, `1.2.3` or a build-specific version string configured at build time).

## Remarks

- Use this method for diagnostics, telemetry, and feature gating by runtime version.
- Method names are case-insensitive in VBScript dispatch.

## Example

```asp
<%
Option Explicit
Dim ax, versionText

Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

versionText = ax.AxVersion()
Response.Write "G3Pix AxonASP Version: " & versionText

Set ax = Nothing
%>
```

## API Reference

- **Object**: `G3AXON.FUNCTIONS`
- **Method**: `AxVersion`
- **Arguments**: none
- **Returns**: `String` (current runtime version)


