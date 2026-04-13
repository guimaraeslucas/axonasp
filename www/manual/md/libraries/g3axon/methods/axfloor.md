# Round Down a Number

Rounds a number down to the next lowest integer.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
Number (Double) = obj.AxFloor(value)
```

## Return Value
Returns a Number (Double).

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxFloor(value)

Response.Write result

Set obj = Nothing
```
