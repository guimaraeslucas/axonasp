# Axuserconfigdirpath

## Overview

Returns the resolved full path to the AxonASP configuration file: config/axonasp.toml.

## Syntax

```asp
result = obj.Axuserconfigdirpath()
```

## Return Values

- Returns a String.


## Code Example

```asp
<%
Option Explicit
Dim obj, cfg
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
cfg = obj.Axuserconfigdirpath()
Response.Write Server.HTMLEncode(cfg)
Set obj = Nothing
%>
```