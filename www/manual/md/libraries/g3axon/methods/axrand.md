# Generate a Random Number

Generates a random number between the specified minimum and maximum values.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
Number (Double) = obj.AxRand(min, max)
```

## Return Value
Returns a Number (Double).

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxRand(min, max)

Response.Write result

Set obj = Nothing
```
