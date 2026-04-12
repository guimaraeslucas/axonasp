<%@ Language=VBScript %>
<html>
<head>
    <title>Runtime Checklist Test</title>
    <meta charset="utf-8">
</head>
<body>
<%
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")

Sub WriteCheck(name, passed, details)
    If passed Then
        Response.Write "PASS - " & name
    Else
        Response.Write "FAIL - " & name
    End If
    If details <> "" Then
        Response.Write " : " & Server.HTMLEncode(details)
    End If
    Response.Write "<br>"
End Sub

Function ExitFunctionSample()
    ExitFunctionSample = "before"
    Exit Function
    ExitFunctionSample = "after"
End Function

Dim exitSubResult
exitSubResult = ""
Sub ExitSubSample()
    exitSubResult = "before"
    Exit Sub
    exitSubResult = "after"
End Sub

Dim root, fileA, fileB, fileC, renamedFile
Dim folderOld, folderNew
root = Server.MapPath("/temp/runtime_checklist")
fileA = root & "\\a.txt"
fileB = root & "\\b.txt"
fileC = root & "\\c.txt"
renamedFile = root & "\\renamed.txt"
folderOld = root & "\\oldfolder"
folderNew = root & "\\newfolder"

If fso.FolderExists(root) Then
    fso.DeleteFolder root
End If
fso.CreateFolder root

Dim ts
Set ts = fso.CreateTextFile(fileB, True)
ts.Write "b"
ts.Close
Set ts = fso.CreateTextFile(fileA, True)
ts.Write "a"
ts.Close
Set ts = fso.CreateTextFile(fileC, True)
ts.Write "c"
ts.Close

Dim folderObj, fileObj, order, item
Set folderObj = fso.GetFolder(root)
order = ""
For Each item In folderObj.Files
    If order <> "" Then order = order & ","
    order = order & LCase(item.Name)
Next
Call WriteCheck("GetFolder files alphabetical", order = "a.txt,b.txt,c.txt", "order=" & order)

Set fileObj = fso.GetFile(fileA)
fileObj.Name = "renamed.txt"
Call WriteCheck("FSO file rename via Name property", fso.FileExists(renamedFile) And (Not fso.FileExists(fileA)), "renamed=" & CStr(fso.FileExists(renamedFile)))

fso.CreateFolder folderOld
Set folderObj = fso.GetFolder(folderOld)
folderObj.Name = "newfolder"
Call WriteCheck("FSO folder rename via Name property", fso.FolderExists(folderNew) And (Not fso.FolderExists(folderOld)), "renamed=" & CStr(fso.FolderExists(folderNew)))

Call WriteCheck("Exit Function", ExitFunctionSample() = "before", "value=" & CStr(ExitFunctionSample()))
Call ExitSubSample()
Call WriteCheck("Exit Sub", exitSubResult = "before", "value=" & exitSubResult)

Application("appA") = "1"
Application("appB") = "2"
Application.Contents.RemoveAll
Call WriteCheck("Application.Contents.RemoveAll", Application.Contents.Count = 0, "count=" & CStr(Application.Contents.Count))

Session("sessA") = "1"
Session("sessB") = "2"
Session.Contents.RemoveAll
Call WriteCheck("Session.Contents.RemoveAll", Session.Contents.Count = 0, "count=" & CStr(Session.Contents.Count))

Session("sessAfter") = "persist"
Session.Abandon
Call WriteCheck("Session.Abandon clears current session", Session.Contents.Count = 0, "count=" & CStr(Session.Contents.Count))

If fso.FolderExists(root) Then
    fso.DeleteFolder root
End If
%>
</body>
</html>
