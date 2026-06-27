<%
@ Language = VBScript
%>
<%
Option Explicit

Dim mode
mode = LCase(Trim(CStr(Request.QueryString("mode"))))

If mode = "image" Then
    Call RenderJpegBinary()
Else
    Call RunTests()
End If

Sub RenderJpegBinary()
    On Error Resume Next

    Dim jpeg, bytes
    Set jpeg = Server.CreateObject("Persits.Jpeg")

    ' Create 400x200 canvas
    jpeg.New 400, 200, "#0F172A"

    ' Set pen properties and draw shapes
    jpeg.Canvas.Pen.Color = &HFF0000 ' Red
    jpeg.Canvas.Pen.Width = 3
    jpeg.Canvas.DrawLine 10, 10, 390, 10
    jpeg.Canvas.DrawLine 10, 190, 390, 190

    ' Draw a filled bar/rectangle using Pen.Color
    jpeg.Canvas.Pen.Color = &H14B8A6 ' Teal
    jpeg.Canvas.DrawBar 50, 40, 350, 80

    ' Set font properties and draw text
    jpeg.Canvas.Font.Family = "Arial"
    jpeg.Canvas.Font.Size = 16
    jpeg.Canvas.Font.Color = &HFFFFFF ' White
    jpeg.Canvas.Font.Bold = True
    jpeg.Canvas.PrintText 100, 130, "Persits.Jpeg Compat Test"

    bytes = jpeg.SendBinary()
    jpeg.Close
    Set jpeg = Nothing

    If Err.Number <> 0 Then
        Response.ContentType = "text/plain"
        Response.Write "Persits.Jpeg render failed: " & Err.Description
        Response.End
    End If

    Response.ContentType = "image/jpeg"
    Response.BinaryWrite bytes
    Response.End
End Sub

Sub RunTests()
    On Error Resume Next
    Dim outText, jpeg, passed, canvas, font, pen, tempPath

    passed = True
    outText = "<h2>Persits.Jpeg Compatibility Verification</h2>"

    Set jpeg = Server.CreateObject("Persits.Jpeg")
    If Err.Number <> 0 Then
        outText = outText & "<p class='bad'>Failed to create object Persits.Jpeg: " & Err.Description & "</p>"
        passed = False
    Else
        outText = outText & "<p class='ok'>Successfully instantiated Persits.Jpeg</p>"
    End If

    ' Test New canvas
    jpeg.New 200, 100, &HFFFFFF
    If jpeg.Width <> 200 Or jpeg.Height <> 100 Then
        outText = outText & "<p class='bad'>Failed: dimensions are not 200x100 (got " & jpeg.Width & "x" & jpeg.Height & ")</p>"
        passed = False
    Else
        outText = outText & "<p class='ok'>New Canvas dimensions verified: " & jpeg.Width & "x" & jpeg.Height & "</p>"
    End If

    ' Verify original dimensions
    If jpeg.OriginalWidth <> 200 Or jpeg.OriginalHeight <> 100 Then
        outText = outText & "<p class='bad'>Failed: original dimensions incorrect (got " & jpeg.OriginalWidth & "x" & jpeg.OriginalHeight & ")</p>"
        passed = False
    Else
        outText = outText & "<p class='ok'>Original dimensions verified: " & jpeg.OriginalWidth & "x" & jpeg.OriginalHeight & "</p>"
    End If

    ' Verify properties get/set
    jpeg.Quality = 85
    If jpeg.Quality <> 85 Then
        outText = outText & "<p class='bad'>Failed: Quality get/set (got " & jpeg.Quality & ")</p>"
        passed = False
    Else
        outText = outText & "<p class='ok'>Quality verified: " & jpeg.Quality & "</p>"
    End If

    jpeg.Interpolation = 3
    If jpeg.Interpolation <> 3 Then
        outText = outText & "<p class='bad'>Failed: Interpolation get/set (got " & jpeg.Interpolation & ")</p>"
        passed = False
    Else
        outText = outText & "<p class='ok'>Interpolation verified: " & jpeg.Interpolation & "</p>"
    End If

    ' Verify resizing properties
    jpeg.Width = 100
    If jpeg.Width <> 100 Then
        outText = outText & "<p class='bad'>Failed: Width resizing (got " & jpeg.Width & ")</p>"
        passed = False
    Else
        outText = outText & "<p class='ok'>Width resizing verified: " & jpeg.Width & "</p>"
    End If

    ' Verify Pen sub-object
    Set pen = jpeg.Canvas.Pen
    pen.Color = &H0000FF ' Blue
    pen.Width = 2.5
    If pen.Width <> 2.5 Then
        outText = outText & "<p class='bad'>Failed: Pen.Width (got " & pen.Width & ")</p>"
        passed = False
    Else
        outText = outText & "<p class='ok'>Pen.Width verified: " & pen.Width & "</p>"
    End If

    ' Verify Font sub-object
    Set font = jpeg.Canvas.Font
    font.Family = "Arial"
    font.Size = 14
    font.Bold = True
    font.Italic = True
    font.Color = &HFF0000 ' Red
    If font.Size <> 14 Or font.Bold <> True Then
        outText = outText & "<p class='bad'>Failed: Font.Size/Bold (got " & font.Size & ", bold=" & font.Bold & ")</p>"
        passed = False
    Else
        outText = outText & "<p class='ok'>Font properties verified: Size=" & font.Size & ", Bold=" & font.Bold & ", Italic=" & font.Italic & "</p>"
    End If

    ' Test Save to temp
    tempPath = Server.MapPath("temp/test_save_persits.jpg")
    jpeg.Save tempPath
    If Err.Number <> 0 Then
        outText = outText & "<p class='bad'>Failed to Save: " & Err.Description & "</p>"
        passed = False
    Else
        outText = outText & "<p class='ok'>Successfully saved image to " & tempPath & "</p>"
    End If

    ' Test Open of saved file
    Dim jpeg2
    Set jpeg2 = Server.CreateObject("Persits.Jpeg")
    jpeg2.Open tempPath
    If Err.Number <> 0 Then
        outText = outText & "<p class='bad'>Failed to Open saved file: " & Err.Description & "</p>"
        passed = False
    Else
        outText = outText & "<p class='ok'>Successfully opened saved image: " & jpeg2.Width & "x" & jpeg2.Height & "</p>"
    End If

    ' Clean up
    jpeg.Close
    jpeg2.Close
    Set jpeg = Nothing
    Set jpeg2 = Nothing

    ' Render output HTML page
%>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>Persits.Jpeg Compatibility Test</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                padding: 24px;
                line-height: 1.45;
            }
            .box {
                border: 1px solid #ddd;
                border-radius: 8px;
                padding: 14px;
                margin-bottom: 16px;
            }
            .ok {
                color: #0f766e;
            }
            .bad {
                color: #b91c1c;
            }
            img {
                max-width: 100%;
                border: 1px solid #ddd;
                border-radius: 8px;
            }
        </style>
    </head>
    <body>
        <h1>Persits.Jpeg ASP Test</h1>
        <div class="box">
            <%= outText %>
            <% If passed Then %>
                <h3 class="ok">STATUS: ALL TESTS PASSED</h3>
            <% Else %>
                <h3 class="bad">STATUS: TEST FAILURE</h3>
            <% End If %>
        </div>

        <div class="box">
            <h3>Rendered Canvas Output</h3>
            <img src="test_persits_jpeg.asp?mode=image" alt="persits output" />
        </div>
    </body>
</html>
<%
End Sub
%>
