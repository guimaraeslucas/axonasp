# Use the G3IMAGE Library

## Overview
The **G3IMAGE** library is a high-performance native image processing and 2D drawing engine for G3Pix AxonASP. It enables developers to programmatically create, load, manipulate, and render images directly on the server. The library supports standard formats including PNG and JPEG, and provides a rich set of drawing primitives such as lines, rectangles, circles, and high-quality text rendering with TrueType font support.

## Syntax
To instantiate the library, use the following syntax:
```asp
Set image = Server.CreateObject("G3IMAGE")
```

## Prerequisites
- **Memory**: Ensure sufficient server memory for high-resolution image processing.
- **File System Permissions**: Requires read/write access to the directories where images are loaded or saved.
- **Fonts**: TrueType (.ttf) or OpenType (.otf) font files are required for advanced text rendering.

## How it Works
The G3IMAGE object operates by managing a drawing context. You can either create a new blank context using **NewContext** or load an existing image using **LoadImage**. Once a context is established, you can perform drawing operations like **DrawLine**, **DrawRectangle**, and **DrawString**. 

After completing the drawing operations, the image can be saved to disk using **SavePNG** or **SaveJPG**, or rendered directly to a byte array using **RenderViaTemp** for immediate delivery to the client via `Response.BinaryWrite`.

The library also supports advanced features like anti-aliasing, alpha transparency (PNG), and configurable JPEG compression quality.

## API Reference

### Methods
- **Close**: Releases all internal resources and clears the active context.
- **DrawCircle**: Draws a circle at the specified coordinates.
- **DrawEllipse**: Draws an ellipse at the specified coordinates.
- **DrawImage**: Draws a previously loaded image onto the current canvas.
- **DrawLine**: Draws a straight line between two points.
- **DrawRectangle**: Draws a rectangle with the specified dimensions.
- **DrawString**: Renders text at the specified coordinates.
- **DrawStringAnchored**: Renders text with precise anchor alignment.
- **Fill**: Fills the current path with the active color.
- **LoadFontFace**: Loads a TrueType font for text rendering.
- **LoadImage**: Loads an image file into memory.
- **LoadJPG**: Explicitly loads a JPEG image file.
- **LoadPNG**: Explicitly loads a PNG image file.
- **MeasureString**: Returns the width and height of a text string in pixels.
- **NewContext**: Initializes a new drawing canvas with the specified width and height.
- **RenderViaTemp**: Renders the current image to a byte array for binary output.
- **SaveJPG**: Saves the current canvas to a JPEG file.
- **SavePNG**: Saves the current canvas to a PNG file.
- **SetColor**: Sets the active drawing color using RGB or RGBA values.
- **SetHexColor**: Sets the active drawing color using a hexadecimal string.
- **SetLineWidth**: Configures the thickness of lines and strokes.
- **Stroke**: Outlines the current path with the active color.

### Properties
- **DefaultFormat**: Gets or sets the preferred output format (png or jpg).
- **HasContext**: Indicates whether a drawing canvas is currently initialized.
- **Height**: Returns the height of the current canvas in pixels.
- **JPGQuality**: Gets or sets the compression quality for JPEG output (1-100).
- **LastBytes**: Returns the raw byte array from the last render operation.
- **LastError**: Returns the description of the most recent error.
- **LastMimeType**: Returns the MIME type of the last rendered image.
- **LastTempFile**: Returns the path to the temporary file used during rendering.
- **Width**: Returns the width of the current canvas in pixels.

## Code Example
The following example demonstrates how to create a simple image with text and output it to the browser.

```asp
<%
Dim img, bytes
Set img = Server.CreateObject("G3IMAGE")

' Create a 400x200 canvas
If img.NewContext(400, 200) Then
    ' Fill background with white
    img.SetHexColor "#FFFFFF"
    img.Clear
    
    ' Draw a blue rectangle
    img.SetHexColor "#003399"
    img.DrawRectangle 10, 10, 380, 180
    img.SetLineWidth 2
    img.Stroke
    
    ' Draw some text
    img.SetHexColor "#333333"
    ' Note: Requires a valid font path
    If img.LoadFontFace("C:\Windows\Fonts\arial.ttf", 24) Then
        img.DrawStringAnchored "AxonASP G3IMAGE", 200, 100, 0.5, 0.5
    End If
    
    ' Render to byte array as PNG
    bytes = img.RenderViaTemp("png")
    
    ' Output to browser
    Response.ContentType = img.LastMimeType
    Response.BinaryWrite bytes
Else
    Response.Write "Error: " & img.LastError
End If

Set img = Nothing
%>
```
