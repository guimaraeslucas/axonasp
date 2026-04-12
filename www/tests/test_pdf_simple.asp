<%
@ Language = VBScript
%>
<%
Option Explicit

On Error Resume Next

Dim pdf
Set pdf = Server.CreateObject("G3PDF")

If Err.Number <> 0 Then
    Response.Write "Error creating PDF: " & Err.Description
    Response.End
End If

pdf.AddPage
pdf.SetFont "helvetica", "", 12

pdf.Write 5, "Test paragraph 1"
pdf.Ln 3
pdf.Write 5, "Test paragraph 2"

Response.ContentType = "application/pdf"
Call pdf.Output("I", "test.pdf", True)

If Err.Number <> 0 Then
    Response.Write "Error in Output: " & Err.Description
    Response.End
End If
%>

