# G3IMAGE Methods

## Overview
This page provides a summary of the methods available in the **G3IMAGE** library for image processing and 2D drawing within the AxonASP environment.

## Method List

- **Close**: Releases all active image contexts and memory buffers.
- **DrawCircle**: Draws a circular outline or path on the current canvas.
- **DrawEllipse**: Draws an elliptical outline or path on the current canvas.
- **DrawImage**: Overlays a loaded image onto the current drawing context.
- **DrawLine**: Draws a line segment between two specified coordinate points.
- **DrawRectangle**: Draws a rectangular outline or path on the current canvas.
- **DrawString**: Renders a text string at a specific position on the canvas.
- **DrawStringAnchored**: Renders text with precise horizontal and vertical anchor points.
- **Fill**: Fills the interior of the current drawing path with the active color.
- **LoadFontFace**: Loads a TrueType or OpenType font for subsequent text rendering.
- **LoadImage**: Loads an image from a file path into the internal image buffer.
- **LoadJPG**: Decodes a JPEG image file from a specified path.
- **LoadPNG**: Decodes a PNG image file from a specified path.
- **MeasureString**: Calculates the pixel dimensions (width and height) of a text string.
- **NewContext**: Creates a blank image canvas with specified dimensions.
- **RenderViaTemp**: Renders the current canvas to a byte array using a temporary file.
- **SaveJPG**: Encodes the current canvas as a JPEG file and saves it to disk.
- **SavePNG**: Encodes the current canvas as a PNG file and saves it to disk.
- **SetColor**: Sets the active drawing color using RGB or RGBA components.
- **SetHexColor**: Sets the active drawing color using a hexadecimal color code.
- **SetLineWidth**: Configures the thickness of subsequent strokes and outlines.
- **Stroke**: Outlines the current drawing path with the active color.

## Remarks
- Method names are case-insensitive.
- Most drawing methods require an active context initialized via **NewContext** or **ContextForImage**.
- Geometric methods like **DrawLine** define paths that must be followed by **Stroke** or **Fill** to appear.
