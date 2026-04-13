# G3IMAGE Properties

## Overview
This page provides a summary of the properties available in the **G3IMAGE** library for inspecting canvas state and configuring rendering options.

## Property List

- **DefaultFormat**: Read/Write. Gets or sets the preferred output format (png or jpg) for rendering.
- **HasContext**: Read-only. Returns a **Boolean** indicating if a drawing canvas is initialized.
- **Height**: Read-only. Returns the pixel height of the current drawing canvas.
- **JPGQuality**: Read/Write. Gets or sets the quality level (1-100) for JPEG encoding.
- **LastBytes**: Read-only. Returns the raw byte array from the most recent render operation.
- **LastError**: Read-only. Returns the error message string from the most recent operation.
- **LastMimeType**: Read-only. Returns the MIME type (e.g., image/png) of the last rendered image.
- **LastTempFile**: Read-only. Returns the path of the temporary file generated during rendering.
- **Width**: Read-only. Returns the pixel width of the current drawing canvas.

## Remarks
- Read-only properties will raise a runtime error if an assignment is attempted.
- Property names are case-insensitive.
- Accessing properties is efficient and does not trigger expensive image processing.
