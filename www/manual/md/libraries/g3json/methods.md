# G3JSON Methods

## Overview
This page provides a summary of the methods available in the **G3JSON** library for JSON processing within the AxonASP environment.

## Method List

- **LoadFile**: Reads a JSON file from the disk and parses it into a native object or array.
- **NewArray**: Creates a new, empty VBScript-compatible array.
- **NewObject**: Creates a new, empty **Scripting.Dictionary** object.
- **Parse**: Converts a JSON-formatted string into a native AxonASP structure.
- **Stringify**: Serializes an AxonASP structure (Dictionary, Array, or primitive) into a JSON string.

## Remarks
- Method names are case-insensitive.
- Parsing returns a **Scripting.Dictionary** object for JSON objects and a standard **Array** for JSON arrays.
- Stringification supports nested objects and arrays.
