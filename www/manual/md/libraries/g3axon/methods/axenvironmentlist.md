# axenvironmentlist

## Overview
Retrieves a list of all environment variables for the current G3Pix AxonASP process.

## Syntax
```asp
result = obj.axenvironmentlist()
```

## Parameters and Arguments
None.

## Return Values
Returns an Array of Strings, where each element represents an environment variable in "KEY=VALUE" format.

## Remarks
Internal or pseudo-environment variables (such as those starting with '=' on Windows) are automatically filtered from the results.

## Code Example
```asp
<%
Option Explicit
Dim obj, envList, item
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

envList = obj.axenvironmentlist()

If IsArray(envList) Then
    For Each item In envList
        Response.Write item & "<br>"
    Next
End If

Set obj = Nothing
%>
```
