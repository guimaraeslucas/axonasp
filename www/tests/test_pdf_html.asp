<%@ Language=VBScript %>
<%
Option Explicit

Dim pdf, htmlPath
Set pdf = Server.CreateObject("G3PDF")

htmlPath = Server.MapPath("/tests/test_pdf_html_sample.html")

pdf.AddPage
pdf.SetTitle "PDF HTML Test", True
pdf.SetFont "helvetica", "", 11
pdf.WriteHTMLFile htmlPath

Response.ContentType = "application/pdf"
pdf.Output "I", "test_pdf_html.pdf", True
%>
