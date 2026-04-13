# Close Method

## Overview
Concludes active session streams alongside removing pending blockades on referenced TAR storage items safely.

## Syntax
```asp
Dim success
success = obj.Close()
```

## Parameters and Arguments
- Close invocation relies entirely on internal tracking; no parameters required.

## Return Values
Signals operational completion reliably transmitting a `Boolean` equating True indicating termination of locks handled nicely.

## Remarks
- Necessary logic assuring background OS file permissions get properly disposed.
- Abandoned locks could halt consecutive actions targeting similar volumes producing unexpected errors.

## Code Example
```asp
<%
Option Explicit
Dim obj
Set obj = Server.CreateObject("G3TAR")
If obj.Open("C:\temp\test.tar") Then
    ' File lock assumed...
    obj.Close()
End If
Set obj = Nothing
%>
```