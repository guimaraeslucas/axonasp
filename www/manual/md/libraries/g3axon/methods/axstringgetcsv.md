# Parse CSV String

## Overview

Parses a comma-separated values (CSV) string and returns an array of values.

## Syntax

```vbscript
arrayResult = obj.axstringgetcsv(str[, delimiter])
```

## Parameters

- **str** (String): The CSV string to parse.
- **delimiter** (String, Optional): The delimiter character. Defaults to `,`.

## Return Value

Variant (Array of Strings). Contains the parsed values from the CSV string.

## Remarks

If the string is empty or parsing fails, an empty array is returned.

## Code Example

```vbscript
Dim obj, arr, str
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
str = "apple,banana,orange"
arr = obj.axstringgetcsv(str)
Response.Write arr(0) ' Outputs: apple
```