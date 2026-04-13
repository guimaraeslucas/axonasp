# NewObject Method

## Overview
Creates a new, empty **Scripting.Dictionary** object.

## Syntax
```asp
Set dict = json.NewObject()
```

## Return Values
Returns a **Scripting.Dictionary** object handle.

## Remarks
This method is a helper for creating native dictionary objects that are compatible with the **Stringify** method. It is functionally equivalent to `Server.CreateObject("Scripting.Dictionary")`.

## Code Example
```asp
<%
Dim json, dict
Set json = Server.CreateObject("G3JSON")
Set dict = json.NewObject()
dict.Add "key", "value"
Response.Write "Dictionary created and populated"
Set json = Nothing
%>
```
