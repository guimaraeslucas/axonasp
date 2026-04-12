<%
@ Language = VBScript
%>
<%
Option Explicit

Dim pdf, logoPath, html
Set pdf = Server.CreateObject("G3PDF")

logoPath = Replace(Server.MapPath("/tests/logo.png"), "\\", "/")
html = "<h1>AxonASP PDF HTML Test</h1>" & _
       "<p>This file validates the HTML renderer integrated with G3PDF.</p>" & _
       "<p>This is the first paragraph with enough text to validate line height, wrapping, and spacing before the next paragraph.</p>" & _
       "<p>This is the second paragraph. It must start below the previous one without any text overlap.</p>" & _
       "<p><b>Bold</b>, <i>Italic</i>, <u>Underline</u>, and a " & _
       "<a href=" & Chr(34) & "https://codeberg.org/go-pdf/fpdf" & Chr(34) & ">link</a>.</p>" & _
       "<p><span style=" & Chr(34) & "color:#cc0000" & Chr(34) & ">Red text</span> and " & _
       "<span style=" & Chr(34) & "color:#0055aa" & Chr(34) & ">blue text</span>.</p>" & _
       "<hr>" & _
       "<table>" & _
       "<tr>" & _
       "<th width=" & Chr(34) & "45" & Chr(34) & ">Column</th>" & _
       "<th width=" & Chr(34) & "55" & Chr(34) & ">Value</th>" & _
       "<th width=" & Chr(34) & "70" & Chr(34) & ">Notes</th>" & _
       "</tr>" & _
       "<tr>" & _
       "<td style=" & Chr(34) & "background-color:#f4cccc; color:#990000" & Chr(34) & ">Background</td>" & _
       "<td style=" & Chr(34) & "background-color:#d9ead3; color:#274e13" & Chr(34) & ">Text color</td>" & _
       "<td style=" & Chr(34) & "background-color:#cfe2f3; color:#073763" & Chr(34) & ">Basic HTML table rendering</td>" & _
       "</tr>" & _
       "</table>" & _
       "<p>Embedded image from /www/tests/logo.png:</p>" & _
       "<img src=" & Chr(34) & logoPath & Chr(34) & " width=" & Chr(34) & "36" & Chr(34) & ">"

pdf.AddPage
pdf.SetTitle "PDF HTML Test", False
pdf.SetFont "helvetica", "", 11
pdf.WriteHTML html

Response.ContentType = "application/pdf"
pdf.Output "I", "test_pdf_html.pdf", True
%>
