# Get Maximum Integer Values

Returns the maximum integer value supported by the system.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
Number (Long) = obj.AxIntegerMax()
```

## Return Value
Returns a Number (Long).

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxIntegerMax()

Response.Write result

Set obj = Nothing
```
