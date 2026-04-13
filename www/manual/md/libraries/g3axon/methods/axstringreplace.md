# Replace Substrings in a String

Replaces all occurrences of the search string with the replacement string.

## Prerequisites
The `G3AXON.FUNCTIONS` object must be instantiated to use this method.
This feature is available in the G3Pix AxonASP environment.

## Syntax
```vbscript
String = obj.AxStringReplace(search, replace, subject)
```

## Return Value
Returns a String.

## Example
```vbscript
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.AxStringReplace(search, replace, subject)

Response.Write result

Set obj = Nothing
```
