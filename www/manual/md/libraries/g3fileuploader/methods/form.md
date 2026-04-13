# Form Method

## Overview
Equivalent to `Request.Form("key")` but safely integrated with multipart form data upload flows. Given that normal request processing may be trapped by stream bodies, this extracts standard form field items alongside file uploads. The alias `FormValue` is also supported.

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
Dim description
description = uploader.Form("fileDescription")
```

## Parameters and Arguments
- `FieldName` (String, Required): The HTML element `name` variable of the standard form field.

## Return Values
Returns a string representing the extracted value, or `Empty` if not found or if form analysis fails completely.
