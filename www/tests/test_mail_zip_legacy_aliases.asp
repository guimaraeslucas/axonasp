<%@ Language=VBScript %>
<%
Response.ContentType = "text/plain"
On Error Resume Next

Sub PrintStatus(label)
    Response.Write label & "|Err=" & Err.Number & "|Desc=" & Err.Description & vbCrLf
    Err.Clear
End Sub

Dim obj

Set obj = Server.CreateObject("Persits.MailSender")
If IsObject(obj) Then
    obj.From = "from@example.com"
    obj.Subject = "Legacy test"
    obj.Body = "Body"
    obj.AddAddress "to@example.com"
End If
Call PrintStatus("Persits.MailSender")

Set obj = Server.CreateObject("CDO.Message")
If IsObject(obj) Then
    obj.From = "from@example.com"
    obj.To = "to@example.com"
    obj.Subject = "Legacy test"
    obj.TextBody = "Body"
End If
Call PrintStatus("CDO.Message")

Set obj = Server.CreateObject("CDONTS.NewMail")
If IsObject(obj) Then
    obj.From = "from@example.com"
    obj.To = "to@example.com"
    obj.Subject = "Legacy test"
    obj.Body = "Body"
    obj.BodyFormat = 1
End If
Call PrintStatus("CDONTS.NewMail")

Set obj = Server.CreateObject("XStandard.Zip")
If IsObject(obj) Then
    Dim path1
    path1 = "temp/legacy-xstandard.zip"
    obj.Create path1
    obj.AddText "test.txt", "legacy"
    obj.Close
End If
Call PrintStatus("XStandard.Zip")

Set obj = Server.CreateObject("ASPZip.EasyZip")
If IsObject(obj) Then
    Dim path2
    path2 = "temp/legacy-aspzip.zip"
    obj.Create path2
    obj.AddText "test.txt", "legacy"
    obj.Close
End If
Call PrintStatus("ASPZip.EasyZip")
%>
