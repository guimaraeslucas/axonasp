# JpgQuality Property

## Overview
Gets or sets the quality level for JPEG encoding in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
' Get current quality
q = obj.JpgQuality

' Set new quality
obj.JpgQuality = 85
```

## Return Values
Returns an Integer representing the quality level (1 to 100).

## Remarks
- The default quality is 90.
- Higher values result in better image quality but larger file sizes.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
img.JpgQuality = 70
' Perform save or render operations
Set img = Nothing
%>
```
