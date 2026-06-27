# Canvas Property

## Overview
Returns a nested sub-object containing methods and properties to draw text and shapes. This property is part of the `Persits.Jpeg` compatibility layer of the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
Set canvas = obj.Canvas
```

## Remarks
Through `Canvas`, you can traverse the object tree to configure pen and font properties and draw on the image context.

### Canvas Sub-Objects

- **`Canvas.Pen` Properties:**
  - `Color` (Read/Write, String/Integer): The color used for drawing lines and bars. Can be a hex string (e.g., `"#FF0000"`) or a BGR integer.
  - `Width` (Read/Write, Double): The line thickness.

- **`Canvas.Font` Properties:**
  - `Color` (Read/Write, String/Integer): The color used for drawing text.
  - **Family** (Read/Write, String): The font family name (e.g. `"Arial"`) or path to a TrueType font (`.ttf`) file.
  - `Size` (Read/Write, Double): Font size in points.
  - `Bold` (Read/Write, Boolean): Applies bold style.
  - `Italic` (Read/Write, Boolean): Applies italic style.

### Canvas Methods

- `PrintText(x, y, text)`: Draws text at coordinates using the active `Font` properties.
- `DrawLine(x1, y1, x2, y2)`: Draws a line between two coordinates using the active `Pen` properties.
- `DrawBar(x1, y1, x2, y2)`: Draws a solid, filled rectangle using the active `Pen` color.

## Code Example
```asp
<%
Dim jpeg
Set jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.New 400, 300, "#FFFFFF"

' Traversal and properties setting
jpeg.Canvas.Pen.Color = &HFF0000 ' BGR Red
jpeg.Canvas.Pen.Width = 4.0

' Draw a line
jpeg.Canvas.DrawLine 0, 0, 400, 300

' Configure font
jpeg.Canvas.Font.Family = "Arial"
jpeg.Canvas.Font.Size = 18
jpeg.Canvas.Font.Bold = True
jpeg.Canvas.Font.Color = "#00FF00" ' Green

' Print text
jpeg.Canvas.PrintText 50, 150, "Hello World"

jpeg.Save Server.MapPath("canvas_drawing.jpg")
jpeg.Close
Set jpeg = Nothing
%>
```
