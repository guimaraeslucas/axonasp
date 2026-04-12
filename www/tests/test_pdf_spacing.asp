<%
@ Language = VBScript
%>
<%
Option Explicit

Dim pdf, logoPath, html
Set pdf = Server.CreateObject("G3PDF")

logoPath = Replace(Server.MapPath("/tests/logo.png"), "\\", "/")

html = "<h1>AxonASP PDF Test - Spacing</h1>" & _
       "<p>Paragraph 1: This is the first paragraph with enough text to validate line height, wrapping, and spacing before the next paragraph.</p>" & _
       "<p>Paragraph 2: This is the second paragraph. It should have proper spacing above it.</p>" & _
       "<p>Paragraph 3: Testing multiple paragraphs to ensure spacing is consistent.</p>" & _
       "<h2>Colored Text Section</h2>" & _
       "<p>Here is <span style=" & Chr(34) & "color:#cc0000" & Chr(34) & ">red text</span> and " & _
       "<span style=" & Chr(34) & "color:#0055aa" & Chr(34) & ">blue text</span> in one line.</p>" & _
       "<h2>Table with Colors</h2>" & _
       "<table>" & _
       "<tr>" & _
       "<th>Column</th>" & _
       "<th>Value</th>" & _
       "</tr>" & _
       "<tr>" & _
       "<td style=" & Chr(34) & "background-color:#f4cccc; color:#990000" & Chr(34) & ">Red BG</td>" & _
       "<td style=" & Chr(34) & "background-color:#d9ead3; color:#274e13" & Chr(34) & ">Green BG</td>" & _
       "</tr>" & _
       "</table>"

pdf.AddPage
pdf.SetTitle "PDF Spacing Test", False
pdf.SetFont "helvetica", "", 11
pdf.WriteHTML html

pdf.Output "F", Server.MapPath("/temp/test_output.pdf"), True
Response.Write "PDF saved to: " & Server.MapPath("/temp/test_output.pdf")
%>

