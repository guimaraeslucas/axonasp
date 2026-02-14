<%@ Language=VBScript %>
<%
Option Explicit

Dim pdf
Set pdf = Server.CreateObject("G3PDF")

pdf.AddPage
pdf.SetTitle "PDF Basic Test", True
pdf.SetAuthor "AxonASP Test Suite", True
pdf.SetFont "helvetica", "B", 16
pdf.Cell 0, 12, "AxonASP PDF Basic Test", 0, 1, "L", False, ""
pdf.SetFont "helvetica", "", 11
pdf.Cell 0, 8, "This file validates basic page/text rendering.", 0, 1, "L", False, ""
pdf.Cell 0, 8, "Timestamp: " & CStr(Now()), 0, 1, "L", False, ""

Response.ContentType = "application/pdf"
pdf.Output "I", "test_pdf_basic.pdf", True
%>
