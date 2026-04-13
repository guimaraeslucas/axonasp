# axfiltervalidateip

## Overview

The `axfiltervalidateip` method validates whether a given string is a valid IPv4 or IPv6 address.

## Syntax

```asp
result = obj.axfiltervalidateip(ipAddress)
```

## Parameters and Arguments

- **ipAddress** (String): The string to validate as an IP address.

## Return Values

Returns a Boolean indicating whether the string is a valid IP address. Returns `True` if valid, otherwise `False`.

## Remarks

- This method is part of the G3Pix AxonASP library.
- It supports both IPv4 and IPv6 formats.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, ip1, ip2
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

ip1 = "192.168.1.1"
ip2 = "not-an-ip"

If ax.axfiltervalidateip(ip1) Then
    Response.Write ip1 & " is valid.<br>"
End If

If Not ax.axfiltervalidateip(ip2) Then
    Response.Write ip2 & " is invalid.<br>"
End If

Set ax = Nothing
%>
```
