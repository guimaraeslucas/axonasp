# AxGetConfig

## Overview

Use `AxGetConfig` to read one key from the active G3Pix AxonASP configuration.

## Prerequisites

- Instantiate the library with `Server.CreateObject("G3AXON.FUNCTIONS")`.
- Ensure the runtime can load `config/axonasp.toml`.

## Syntax

```asp
value = obj.AxGetConfig(key)
```

## Parameters

- **key** (String): Fully qualified configuration key, such as `global.golang_memory_limit_mb`.

## Return Value

- **Empty**: Returned when `key` is omitted.
- **Empty**: Returned when `key` is blank.
- **Empty**: Returned when the config file is not loaded.
- **Empty**: Returned when the key does not exist.
- **String**: Returned when the resolved config value is textual.
- **Boolean**: Returned when the resolved config value is true/false.
- **Integer**: Returned when the resolved config value is an integer type.
- **Double**: Returned when the resolved config value is a floating-point type.
- **Array**: Returned when the resolved config value is a list.

## Remarks

- When `global.viper_automatic_env` is enabled, environment variables can override file values.
- Method names are case-insensitive in VBScript dispatch.

## Example

```asp
<%
Option Explicit
Dim ax, memLimit, missingKey

Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

memLimit = ax.AxGetConfig("global.golang_memory_limit_mb")
missingKey = ax.AxGetConfig("global.this_key_does_not_exist")

Response.Write "Memory limit: " & CStr(memLimit) & "<br>"
If IsEmpty(missingKey) Then
    Response.Write "Missing key returned Empty"
End If

Set ax = Nothing
%>
```

## API Reference

- **Object**: `G3AXON.FUNCTIONS`
- **Method**: `AxGetConfig`
- **Arguments**: `key As String`
- **Returns**: `Empty`, `String`, `Boolean`, `Integer`, `Double`, or `Array` based on the resolved key state and value type


