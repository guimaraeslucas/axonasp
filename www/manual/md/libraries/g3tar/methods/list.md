# List Method

## Overview
Generates an enumeration of paths that correspond exactly to the current contents found inside the loaded TAR archive.

## Syntax
```asp
Dim stringArray
stringArray = obj.List()
```

## Parameters and Arguments
- No parameters. Read access must first be set up.

## Return Values
Returns a one-dimensional native array storing String definitions, mapped natively within AxonASP.

## Remarks
- Needs a preceding call referencing an archive via Open.
- Ideal when validating content prior to executing intensive extractions or when generating diagnostic listings.

## Code Example
```asp
<%
Option Explicit
Dim obj, files, i
Set obj = Server.CreateObject("G3TAR")
If obj.Open("C:\data\packages\archive.tar") Then
    files = obj.List()
    For i = LBound(files) To UBound(files)
        Response.Write files(i) & "<br>"
    Next
    obj.Close()
End If
Set obj = Nothing
%>
```