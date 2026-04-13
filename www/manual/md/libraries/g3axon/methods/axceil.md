# Round Up a Number

Rounds a number up to the next highest integer.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
Number (Double) = obj.AxCeil(value)
```

## Return Value
Returns a Number (Double).

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxCeil(value)

Response.Write result

Set obj = Nothing
```
