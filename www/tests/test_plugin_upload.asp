<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/html"
Response.Buffer = False
%>
<!DOCTYPE html>
<html>
<head><title>Upload Test with aspL.plugin</title></head>
<body>
<h1>File Upload Test using aspL.plugin("uploader")</h1>

<%
If Request.TotalBytes > 0 Then
    
    Response.Write "<h2>Processing Upload...</h2>" & vbCrLf
    Response.Flush
    
    On Error Resume Next
    
    ' Load the uploader plugin
    Response.Write "<p>Loading uploader plugin...</p>" & vbCrLf
    Response.Flush
    
    Dim uploader
    Set uploader = aspL.plugin("uploader")
    
    If Err.Number <> 0 Then
        Response.Write "<p style='color:red'>Error loading plugin: " & Err.Description & "</p>" & vbCrLf
        Err.Clear
    Else
        Response.Write "<p style='color:green'>Plugin loaded successfully!</p>" & vbCrLf
    End If
    Response.Flush
    
    ' Call Upload method
    Response.Write "<p>Calling Upload()...</p>" & vbCrLf
    Response.Flush
    
    uploader.Upload()
    
    If Err.Number <> 0 Then
        Response.Write "<p style='color:red'>Error in Upload(): " & Err.Description & "</p>" & vbCrLf
        Err.Clear
    Else
        Response.Write "<p style='color:green'>Upload() completed!</p>" & vbCrLf
    End If
    Response.Flush
    
    ' Show form fields
    Response.Write "<h3>Form Fields:</h3>" & vbCrLf
    Response.Write "<p>FormElements count: " & uploader.FormElements.Count & "</p>" & vbCrLf
    
    Dim key
    For Each key In uploader.FormElements.Keys
        Response.Write "<p>" & key & " = " & uploader.FormElements(key) & "</p>" & vbCrLf
    Next
    Response.Flush
    
    ' Show uploaded files
    Response.Write "<h3>Uploaded Files:</h3>" & vbCrLf
    Response.Write "<p>UploadedFiles count: " & uploader.UploadedFiles.Count & "</p>" & vbCrLf
    
    For Each key In uploader.UploadedFiles.Keys
        Dim f
        Set f = uploader.UploadedFiles(key)
        Response.Write "<p>File: " & f.FileName & " (" & f.Length & " bytes, type: " & f.ContentType & ")</p>" & vbCrLf
    Next
    Response.Flush
    
    ' Save files
    If uploader.UploadedFiles.Count > 0 Then
        Response.Write "<h3>Saving Files...</h3>" & vbCrLf
        Response.Flush
        
        Dim savePath
        savePath = Server.MapPath("/uploads")
        Response.Write "<p>Save path: " & savePath & "</p>" & vbCrLf
        
        uploader.Save(savePath)
        
        If Err.Number <> 0 Then
            Response.Write "<p style='color:red'>Error in Save(): " & Err.Description & "</p>" & vbCrLf
            Err.Clear
        Else
            Response.Write "<p style='color:green'>Files saved!</p>" & vbCrLf
        End If
        Response.Flush
        
        ' Verify files saved
        Response.Write "<h3>Verifying Saved Files:</h3>" & vbCrLf
        For Each key In uploader.UploadedFiles.Keys
            Set f = uploader.UploadedFiles(key)
            Response.Write "<p>" & f.FileName & " - Path: " & f.Path & "</p>" & vbCrLf
            If aspL.fso.FileExists(f.Path) Then
                Response.Write "<p style='color:green'>File exists at: " & f.Path & "</p>" & vbCrLf
            Else
                Response.Write "<p style='color:red'>File NOT found at: " & f.Path & "</p>" & vbCrLf
            End If
        Next
    End If
    
    On Error GoTo 0
    
    Set uploader = Nothing
    
Else
    ' Show upload form
%>
    <h2>Upload a File</h2>
    <form method="POST" enctype="multipart/form-data">
        <p><label>File: <input type="file" name="file1"></label></p>
        <p><label>Description: <input type="text" name="description" value="Test upload"></label></p>
        <p><button type="submit">Upload</button></p>
    </form>
<%
End If

Set aspL = Nothing
%>
</body>
</html>
