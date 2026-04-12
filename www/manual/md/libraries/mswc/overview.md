# Use MSWC Components in AxonASP

## Overview
AxonASP supports a compatibility family for classic MSWC components. Members are dispatched case-insensitively and exposed through native object wrappers.

## Supported ProgIDs
- MSWC.AdRotator
- MSWC.BrowserType
- MSWC.NextLink
- MSWC.ContentRotator
- MSWC.Counters
- MSWC.Tools
- MSWC.MyInfo
- MSWC.PageCounter
- MSWC.PermissionChecker

## Documentation Map
- Methods index: [Methods](methods.md)
- Properties index: [Properties](properties.md)
- Member details: methods and properties subfolders.

## Code Example
```asp
<%
Dim bt, canCookies
Set bt = Server.CreateObject("MSWC.BrowserType")
canCookies = bt.Cookies
Response.Write "Browser=" & bt.Browser & " Cookies=" & CStr(canCookies)
%>
```