# AxGetConfigKeys

## Overview

Use `AxGetConfigKeys` to list every configuration key currently visible to G3Pix AxonASP.

## Prerequisites

- Instantiate the library with `Server.CreateObject("G3AXON.FUNCTIONS")`.
- Ensure the runtime can load `config/axonasp.toml` when you expect non-empty output.

## Syntax

```asp
keys = obj.AxGetConfigKeys()
```

## Parameters

- This method does not require parameters.

## Return Value

- **Array**: Zero-based VBArray containing configuration keys as strings.
- **Array**: Returns an empty zero-length array when no keys are available (for example, when the config file is not loaded).

## Remarks

- Each element is a full key name (for example, `global.golang_memory_limit_mb`).
- Method names are case-insensitive in VBScript dispatch.

## Example

```asp
<%
Option Explicit
Dim ax, keys, i

Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

keys = ax.AxGetConfigKeys()
For i = 0 To UBound(keys)
    Response.Write keys(i) & "<br>"
Next

Set ax = Nothing
%>
```

## API Reference

- **Object**: `G3AXON.FUNCTIONS`
- **Method**: `AxGetConfigKeys`
- **Arguments**: none
- **Returns**: `Array` of `String` keys (zero-based VBArray)


