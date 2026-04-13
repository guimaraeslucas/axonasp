# Axshutdownaxonaspserver

## Overview

Shuts down the AxonASP server process when the shutdown function is enabled in configuration.

## Syntax

```asp
result = obj.Axshutdownaxonaspserver(...)
```

## Parameters and Arguments

- Parameters (): None. This method does not accept any parameters.

## Return Values

- Returns a Boolean.


## Remarks

- Method names are case-insensitive.
- For object return values, use Set when assigning the return value.

## Code Example

```asp
<%
Option Explicit
Dim obj
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
obj.Axshutdownaxonaspserver()
'The servel will shut down immediately after this call.
%>
```


