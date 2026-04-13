# Count Elements in an Array

Counts all elements in an array or object.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
Number (Integer) = obj.AxCount(variable)
```

## Return Value
Returns a Number (Integer).

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxCount(variable)

Response.Write result

Set obj = Nothing
```
