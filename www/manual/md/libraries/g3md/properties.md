# G3MD Properties

## Overview
This page provides a summary of the properties available in the **G3MD** library for configuring the Markdown processor.

## Property List

- **HardWraps**: Read/Write. Determines if soft line breaks in the source should be converted to `<br>` tags. Returns a **Boolean**.
- **Unsafe**: Read/Write. Determines if raw HTML and potentially dangerous links should be rendered. Returns a **Boolean**.

## Remarks
- Property names are case-insensitive.
- Setting these properties affects all subsequent calls to the **Process** method on the same object instance.
- By default, both properties are set to **False**.
