# Join Array Elements into a String

Joins the elements of an array with a specified separator string.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
String = obj.AxImplode(separator, array)
```

## Return Value
Returns a String.

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxImplode(separator, array)

Response.Write result

Set obj = Nothing
```
