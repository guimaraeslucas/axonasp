# Get the Smallest Float Value

Returns the smallest float value greater than zero supported by the system.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
Number (Double) = obj.AxSmallestFloatValue()
```

## Return Value
Returns a Number (Double).

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxSmallestFloatValue()

Response.Write result

Set obj = Nothing
```
