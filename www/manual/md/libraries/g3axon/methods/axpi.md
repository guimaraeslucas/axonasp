# Get the Value of Pi

Returns the mathematical constant Pi.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
Number (Double) = obj.AxPi()
```

## Return Value
Returns a Number (Double).

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxPi()

Response.Write result

Set obj = Nothing
```
