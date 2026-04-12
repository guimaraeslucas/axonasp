# Axshutdownaxonaspserver Method

## Overview

Shuts down the AxonASP server process when the shutdown function is enabled in configuration.

## Syntax

```asp
result = obj.Axshutdownaxonaspserver(...)
```

## Parameters and Arguments

- Parameters (): None. This method does not accept any parameters.

## Return Values

- Returns nothing (Empty) on success. If the shutdown function is disabled or an error occurs, a runtime error is raised.

## Remarks

- Method names are case-insensitive.
- For object return values, use Set when assigning the return value.

## Code Example

```asp
<%
Option Explicit
Dim obj
Set obj = Server.CreateObject("G3AXON.Functions")
obj.Axshutdownaxonaspserver()
'The servel will shut down immediately after this call.
%>
```


