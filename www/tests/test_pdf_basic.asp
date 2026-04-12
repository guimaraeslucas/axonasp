<%
@ Language = VBScript
%>
<%
Option Explicit

Dim pdf
Set pdf = Server.CreateObject("G3PDF")

pdf.AddPage
pdf.SetTitle "PDF Basic Test", False
pdf.SetAuthor "AxonASP Test Suite", False
pdf.SetMargins 15, 20, 15
pdf.SetFont "helvetica", "B", 16
pdf.Cell 0, 12, "AxonASP PDF Basic Test", 0, 1, "L", False, ""
pdf.SetFont "helvetica", "", 11
pdf.Cell 0, 8, "This file validates basic page/text rendering.", 0, 1, "L", False, ""
pdf.Cell 0, 8, "Timestamp: " & CStr(Now()), 0, 1, "L", False, ""
pdf.Ln 3
pdf.SetFillColor 31, 78, 121
pdf.SetTextColor 255, 255, 255
pdf.Cell 60, 10, "Header Cell", 1, 0, "C", True, ""
pdf.SetFillColor 226, 239, 218
pdf.SetTextColor 39, 78, 19
pdf.Cell 60, 10, "Green Fill", 1, 0, "C", True, ""
pdf.SetFillColor 252, 229, 205
pdf.SetTextColor 120, 63, 4
pdf.Cell 60, 10, "Orange Fill", 1, 1, "C", True, ""
pdf.SetTextColor 0, 0, 0
pdf.Ln 4
pdf.SetDrawColor 40, 90, 150
pdf.SetLineWidth 0.4
pdf.Rect 15, 60, 80, 30, "D"
pdf.Text 20, 72, "Draw/Text API smoke test"

Response.ContentType = "application/pdf"
pdf.Output "I", "test_pdf_basic.pdf", True
%>
