## G3IMAGE Library Implementation Summary

### Overview
A comprehensive 2D graphics library has been implemented for AxonASP using the `github.com/fogleman/gg` package. The `G3IMAGE` component exposes drawing, text, transforms, gradients, image loading, and rendering APIs for VBScript while preserving AxonASP execution and path security semantics.

### Files Created/Modified

#### New/Modified Files
1. **`server/image_lib.go`**
   - Complete `G3IMAGE` implementation
   - VBScript-accessible wrapper over `gg.Context`
   - Type bridge for images, patterns, gradients, masks, matrices, points, and fonts
   - In-memory rendering helpers and temp-file fallback

#### Integration
1. **`server/executor_libraries.go`**
   - Added `ImageLibrary` wrapper for `ASPLibrary` compatibility
   - Enables `Set img = Server.CreateObject("G3IMAGE")`
   - Also supports `Server.CreateObject("IMAGE")`

2. **`server/executor.go`**
   - Added `CreateObject` mapping for `G3IMAGE` / `IMAGE`

3. **`go.mod` / `go.sum`**
   - Added dependencies required by `gg`

### Key Features Implemented

✓ **Context Creation & Lifecycle**
  - `NewContext(width, height)`
  - `NewContextForImage(imageOrPath)`
  - `NewContextForRGBA(imageOrPath)`

✓ **Drawing & Path Operations**
  - Full drawing primitives: line, point, rectangle, rounded rectangle, circle, ellipse, arcs, regular polygon
  - Path API: `MoveTo`, `LineTo`, `QuadraticTo`, `CubicTo`, `ClosePath`, `ClearPath`, `NewSubPath`
  - Painting API: `Stroke`, `Fill`, `StrokePreserve`, `FillPreserve`, `SetPixel`, `Clear`

✓ **Text & Font Support**
  - `DrawString`, `DrawStringAnchored`, `DrawStringWrapped`
  - `MeasureString`, `MeasureMultilineString`, `WordWrap`, `FontHeight`
  - `LoadFontFace(path, points)` + `SetFontFace(...)`

✓ **Colors, Stroke & Fill Styles**
  - `SetRGB`, `SetRGBA`, `SetRGB255`, `SetRGBA255`, `SetHexColor`, `SetColor`
  - `SetLineWidth`, `SetLineCap*`, `SetLineJoin*`, `SetDash`, `SetDashOffset`
  - `SetFillRule`, `SetFillRuleWinding`, `SetFillRuleEvenOdd`
  - `SetFillStyle(pattern)`, `SetStrokeStyle(pattern)`

✓ **Gradients & Pattern Objects**
  - `NewLinearGradient(...)`
  - `NewRadialGradient(...)`
  - `NewSolidPattern(color)`
  - `NewSurfacePattern(image, repeatOp)`
  - Gradient color stops via `pattern.AddColorStop(offset, color)`

✓ **Transformations & State Stack**
  - `Identity`, `Translate`, `Scale`, `Rotate`, `Shear`
  - `ScaleAbout`, `RotateAbout`, `ShearAbout`, `TransformPoint`, `InvertY`
  - `Push`, `Pop`

✓ **Clipping & Masking**
  - `Clip`, `ClipPreserve`, `ResetClip`
  - `AsMask`, `SetMask(mask)`, `InvertMask`

✓ **Image I/O and Rendering**
  - Loading: `LoadImage`, `LoadPNG`, `LoadJPG`
  - Saving: `SavePNG`, `SaveJPG`
  - Byte rendering: `EncodePNG`, `EncodeJPG`, `RenderContent`
  - Browser-friendly output with `Response.BinaryWrite`

✓ **In-Memory Rendering with Temp Fallback**
  - Primary path: encode directly into memory buffer
  - Fallback path: save to `temp/images` in executable directory, read bytes, return to caller, and delete temp file

### Architecture

**Class Hierarchy**:
```
Component (interface)
  └─ G3IMAGE
      ├─ gg.Context bridge
      ├─ image load/save/render helpers
      ├─ security-aware path resolution
      └─ dynamic method dispatcher for gg methods
```

**Bridged Value Objects**:
- `GGImageValue` (image wrapper)
- `GGPatternValue` (pattern/gradient wrapper)
- `GGMaskValue` (alpha mask wrapper)
- `GGMatrixValue` (matrix wrapper)
- `GGPointValue` (point wrapper)
- `GGFontFaceValue` (font face wrapper)

**Rendering Flow**:
1. ASP script builds drawing state in `G3IMAGE`
2. `gg` renders to the internal context image
3. Output requested as bytes (`RenderContent` / `EncodePNG` / `EncodeJPG`)
4. Bytes sent with `Response.BinaryWrite`

### Properties

#### Core State
- `HasContext` (Boolean)
- `Width` / `Height` (Number)
- `LastError` (String)
- `LastMimeType` (String)
- `LastTempFile` (String)
- `LastBytes` (Byte Array)
- `DefaultFormat` (`png` / `jpg`)
- `JpgQuality` (1-100)

#### Enum-Like Constants
- Align: `AlignLeft`, `AlignCenter`, `AlignRight`
- Fill rules: `FillRuleWinding`, `FillRuleEvenOdd`
- Line caps: `LineCapRound`, `LineCapButt`, `LineCapSquare`
- Line joins: `LineJoinRound`, `LineJoinBevel`
- Repeat ops: `RepeatBoth`, `RepeatX`, `RepeatY`, `RepeatNone`

### Main Methods

#### Context & Resource Setup
- `NewContext(width, height)`
- `NewContextForImage(imageOrPath)`
- `NewContextForRGBA(imageOrPath)`
- `LoadImage(path)`
- `LoadPNG(path)`
- `LoadJPG(path)`
- `Image()` / `GetImage()`

#### Render Output
- `SavePNG(path)`
- `SaveJPG(path, [quality])`
- `EncodePNG()`
- `EncodeJPG([quality])`
- `RenderContent([format], [quality])`
- `RenderViaTemp([format], [quality])`

#### Utility Methods
- `Radians(degrees)`
- `Degrees(radians)`
- `QuadraticBezier(x0, y0, x1, y1, x2, y2)`
- `CubicBezier(x0, y0, x1, y1, x2, y2, x3, y3)`

### Dynamic gg Method Access

`G3IMAGE` forwards most `gg.Context` APIs dynamically, including:
- Geometry/path methods
- Paint/stroke/fill methods
- Text methods
- Transform methods
- Stack methods (`Push` / `Pop`)
- Clip/mask methods
- Style setup methods

This allows VBScript usage with names such as:
- `DrawRectangle`, `DrawCircle`, `DrawStringWrapped`
- `SetLineWidth`, `SetDash`, `SetHexColor`
- `RotateAbout`, `ScaleAbout`, `TransformPoint`

### Usage Examples

#### Basic PNG Render to Browser (In Memory)
```vbscript
<%
Dim img, bytes
Set img = Server.CreateObject("G3IMAGE")

img.NewContext 600, 300
img.SetRGB 1, 1, 1
img.Clear

img.SetHexColor "#1e40af"
img.DrawRoundedRectangle 20, 20, 560, 260, 18
img.Fill

img.SetHexColor "#ffffff"
img.LoadFontFace "assets/fonts/Inter-Regular.ttf", 28
img.DrawStringAnchored "AxonASP + gg", 300, 155, 0.5, 0.5

bytes = img.RenderContent("png")
Response.ContentType = "image/png"
Response.BinaryWrite bytes
Response.End
%>
```

#### JPG Render with Quality
```vbscript
<%
Dim img, bytes
Set img = Server.CreateObject("IMAGE")

img.NewContext 800, 450
img.SetRGB255 30, 30, 30
img.Clear

img.SetRGB255 240, 200, 60
img.DrawCircle 400, 225, 130
img.Fill

bytes = img.RenderContent("jpg", 85)
Response.ContentType = "image/jpeg"
Response.BinaryWrite bytes
%>
```

#### Gradient + Pattern Usage
```vbscript
<%
Dim img, grad, bytes
Set img = Server.CreateObject("G3IMAGE")

img.NewContext 500, 200
Set grad = img.NewLinearGradient(0, 0, 500, 0)
grad.AddColorStop 0.0, "#06b6d4"
grad.AddColorStop 1.0, "#3b82f6"

img.SetFillStyle grad
img.DrawRectangle 0, 0, 500, 200
img.Fill

bytes = img.EncodePNG()
Response.ContentType = "image/png"
Response.BinaryWrite bytes
%>
```

#### Temp-File Fallback (Explicit)
```vbscript
<%
Dim img, bytes
Set img = Server.CreateObject("G3IMAGE")

img.NewContext 300, 120
img.SetHexColor "#111827"
img.Clear

img.SetHexColor "#10b981"
img.DrawString "Temp fallback", 20, 70

bytes = img.RenderViaTemp("png")
Response.ContentType = img.LastMimeType
Response.BinaryWrite bytes
%>
```

### Security Notes

- File paths are resolved through AxonASP execution context (`Server.MapPath` behavior).
- Root-directory checks prevent path traversal outside the configured web root for regular I/O methods.
- Temporary fallback files are created under executable `temp/images` and removed immediately after reading.

### Error Handling

- Errors are captured in `LastError`.
- Methods return `nil`/`false` on failure depending on operation type.
- `LastMimeType`, `LastBytes`, and `LastTempFile` provide post-render diagnostics.

### Third-Party License Attribution

This implementation uses `github.com/fogleman/gg`, licensed under MIT.
Copyright (c) 2016 Michael Fogleman.
