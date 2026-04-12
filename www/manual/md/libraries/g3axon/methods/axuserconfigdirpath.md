# Axuserconfigdirpath Method

## Overview

Returns the resolved full path to the AxonASP configuration file: config/axonasp.toml.

## Syntax

```asp
result = obj.Axuserconfigdirpath()
```

## Return Values

- String: Full path to config/axonasp.toml.

## Code Example

```asp
<%
Option Explicit
Dim obj, cfg
Set obj = Server.CreateObject("G3AXON.Functions")
cfg = obj.Axuserconfigdirpath()
Response.Write Server.HTMLEncode(cfg)
Set obj = Nothing
%>
```