# Reverse an Array

Returns an array with elements in reverse order.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
Array = obj.AxArrayReverse(array)
```

## Return Value
Returns a Array.

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxArrayReverse(array)

Response.Write result

Set obj = Nothing
```
