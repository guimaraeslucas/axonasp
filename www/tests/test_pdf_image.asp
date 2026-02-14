<%@ Language=VBScript %>
<%
Option Explicit

Dim pdf, imagePath
Set pdf = Server.CreateObject("G3PDF")

imagePath = Server.MapPath("/asplite-test/uploads/Png.png")

pdf.AddPage
pdf.SetTitle "PDF Image Test", True
pdf.SetFont "helvetica", "B", 14
pdf.Cell 0, 10, "AxonASP PDF Image Test", 0, 1, "L", False, ""
pdf.SetFont "helvetica", "", 11
pdf.Cell 0, 8, "Image source: /asplite-test/uploads/Png.png", 0, 1, "L", False, ""
pdf.Image imagePath, 20, 35, 70, 0, "", ""

Response.ContentType = "application/pdf"
pdf.Output "I", "test_pdf_image.pdf", True
%>
