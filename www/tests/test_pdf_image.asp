<%
@ Language = VBScript
%>
<%
Option Explicit

Dim pdf, imagePath, fso
Set pdf = Server.CreateObject("G3PDF")
Set fso = Server.CreateObject("Scripting.FileSystemObject")

imagePath = Server.MapPath("/tests/logo.png")

pdf.AddPage
pdf.SetTitle "PDF Image Test", False
pdf.SetFont "helvetica", "B", 14
pdf.Cell 0, 10, "AxonASP PDF Image Test", 0, 1, "L", False, ""
pdf.SetFont "helvetica", "", 11
pdf.Cell 0, 8, "Image source: /tests/logo.png", 0, 1, "L", False, ""

If fso.FileExists(imagePath) Then
    pdf.Image imagePath, 20, 35, 70, 0, "PNG", ""
Else
    pdf.SetTextColor 220, 0, 0
    pdf.Cell 0, 8, "ERROR: /tests/logo.png not found.", 0, 1, "L", False, ""
End If

Response.ContentType = "application/pdf"
pdf.Output "I", "test_pdf_image.pdf", True
%>
