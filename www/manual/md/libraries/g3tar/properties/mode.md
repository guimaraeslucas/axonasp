# Mode Property

## Overview
Gets the current operation mode of the G3TAR archive.

## Syntax
```asp
Dim currentMode
currentMode = obj.Mode
```

## Parameters and Arguments
- Getter: None.

## Return Values
Returns a `String` representing the archive mode.

## Remarks
- This property is read-only.
- Usually reflects read or write modes depending on initialization.

## Code Example
```asp
<%
Option Explicit
Dim obj, currentMode
Set obj = Server.CreateObject("G3TAR")
If obj.Open("C:\temp\archive.tar") Then
    currentMode = obj.Mode
    Response.Write currentMode
End If
obj.Close()
Set obj = Nothing
%>
```

