# axclearenvironment

## Overview
Removes all environment variables from the current G3Pix AxonASP process environment.

## Syntax
```asp
result = obj.axclearenvironment()
```

## Parameters and Arguments
None.

## Return Values
Returns a Boolean (True) indicating that the environment has been cleared.

## Remarks
This operation is destructive and affects the current process only. Use this function with extreme care as it may cause system utilities or other libraries to fail if they depend on specific environment variables.

## Code Example
```asp
<%
Option Explicit
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

' Caution: This will remove ALL environment variables for the current process
result = obj.axclearenvironment()

If result Then
    Response.Write "Environment variables have been cleared."
End If

Set obj = Nothing
%>
```
