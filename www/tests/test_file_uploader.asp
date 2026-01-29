<%
' All VBScript logic at the top
Option Explicit

Dim isPost, action, uploader, result, uploadDir, results, i, fileInfo
Dim simpleSuccess, simpleError, multipleSuccess, multipleError, infoSuccess, infoError
Dim simpleResultObj, multipleResultsArr, fileInfoObj
Dim statusText, statusClass, htmlOutput

isPost = (Request.ServerVariables("REQUEST_METHOD") = "POST")
action = Request.Form("action")

' Initialize error/success flags
simpleSuccess = False
simpleError = ""
multipleSuccess = False
multipleError = ""
infoSuccess = False
infoError = ""

' SIMPLE UPLOAD HANDLING
If isPost And action = "simple" Then
    Set uploader = Server.CreateObject("G3FileUploader")
    uploader.BlockExtensions "exe,dll,bat,cmd"
    uploader.SetProperty "maxfilesize", 5242880
    
    Set simpleResultObj = uploader.Process("file1", "/uploads")
    
    If simpleResultObj("IsSuccess") Then
        simpleSuccess = True
    Else
        simpleSuccess = False
        simpleError = simpleResultObj("ErrorMessage")
    End If
End If

' MULTIPLE FILES UPLOAD HANDLING
If isPost And action = "multiple" Then
    Set uploader = Server.CreateObject("G3FileUploader")
    uploader.AllowExtensions "jpg,png,gif,pdf,doc,docx,xls,xlsx"
    uploader.SetUseAllowedOnly(True)
    
    Set multipleResultsArr = uploader.ProcessAll("/uploads")
    
    If IsArray(multipleResultsArr) And UBound(multipleResultsArr) >= 0 Then
        multipleSuccess = True
    Else
        multipleSuccess = False
        multipleError = "No files uploaded"
    End If
End If

' FILE INFO HANDLING
If isPost And action = "info" Then
    Set uploader = Server.CreateObject("G3FileUploader")
    uploader.BlockExtensions "exe,dll,bat,cmd,vbs,js,scr"
    
    Set fileInfoObj = uploader.GetFileInfo("infofile")
    
    If fileInfoObj("IsSuccess") = False Then
        infoSuccess = False
        infoError = fileInfoObj("ErrorMessage")
    Else
        infoSuccess = True
    End If
End If

' Generate simple upload HTML
If simpleSuccess Then
    Response.Write "<div class='result success'>"
    Response.Write "<strong>Upload Successful!</strong><br>"
    Response.Write "Original Name: " & simpleResultObj("OriginalFileName") & "<br>"
    Response.Write "New Name: " & simpleResultObj("NewFileName") & "<br>"
    Response.Write "Size: " & simpleResultObj("Size") & " bytes<br>"
    Response.Write "MIME Type: " & simpleResultObj("MimeType") & "<br>"
    Response.Write "Extension: " & simpleResultObj("Extension") & "<br>"
    Response.Write "Relative Path: " & simpleResultObj("RelativePath") & "</div>"
ElseIf simpleError <> "" Then
    Response.Write "<div class='result error'>"
    Response.Write "<strong>Upload Failed!</strong><br>"
    Response.Write "Error: " & simpleError & "</div>"
End If

' Generate multiple upload HTML
If multipleSuccess And IsArray(multipleResultsArr) Then
    Response.Write "<table><tr><th>File Name</th><th>Size (bytes)</th><th>MIME Type</th><th>Extension</th><th>Status</th></tr>"
    For i = 0 To UBound(multipleResultsArr)
        If multipleResultsArr(i)("IsSuccess") Then
            statusText = "OK"
            statusClass = "success"
        Else
            statusText = "FAILED: " & multipleResultsArr(i)("ErrorMessage")
            statusClass = "error"
        End If
        Response.Write "<tr><td>" & multipleResultsArr(i)("OriginalFileName") & "</td>"
        Response.Write "<td>" & multipleResultsArr(i)("Size") & "</td>"
        Response.Write "<td>" & multipleResultsArr(i)("MimeType") & "</td>"
        Response.Write "<td>" & multipleResultsArr(i)("Extension") & "</td>"
        Response.Write "<td class='" & statusClass & "'>" & statusText & "</td></tr>"
    Next
    Response.Write "</table>"
ElseIf multipleError <> "" Then
    Response.Write "<div class='result error'>" & multipleError & "</div>"
End If

' Generate file info HTML
If infoSuccess Then
    Response.Write "<div class='result info'>"
    Response.Write "<strong>File Information:</strong><br>"
    Response.Write "File Name: " & fileInfoObj("OriginalFileName") & "<br>"
    Response.Write "Size: " & fileInfoObj("Size") & " bytes<br>"
    Response.Write "MIME Type: " & fileInfoObj("MimeType") & "<br>"
    Response.Write "Extension: " & fileInfoObj("Extension") & "<br>"
    If fileInfoObj("IsValid") Then
        Response.Write "Is Valid: Yes<br>"
    Else
        Response.Write "Is Valid: No<br>"
    End If
    If fileInfoObj("ExceedsMaxSize") Then
        Response.Write "Exceeds Max Size: Yes"
    Else
        Response.Write "Exceeds Max Size: No"
    End If
    Response.Write "</div>"
ElseIf infoError <> "" Then
    Response.Write "<div class='result error'>" & infoError & "</div>"
End If
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>G3FileUploader - Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .container { max-width: 800px; margin: 0 auto; }
        .test-section { margin-bottom: 30px; padding: 15px; border: 1px solid #ccc; border-radius: 5px; }
        .result { background-color: #f0f0f0; padding: 10px; margin-top: 10px; border-radius: 3px; }
        .error { color: red; }
        .success { color: green; }
        .info { color: blue; }
        form { margin-top: 15px; }
        input[type="file"] { margin: 10px 0; }
        input[type="submit"] { padding: 8px 15px; cursor: pointer; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
        th { background-color: #f0f0f0; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3FileUploader - Upload Test</h1>
        
        <div class="test-section">
            <h2>Simple File Upload</h2>
            <p>Upload a single file to the /uploads directory</p>
            
            <% Response.Write "" %>
            
            <form method="POST" enctype="multipart/form-data">
                <input type="hidden" name="action" value="simple">
                <input type="file" name="file1" required>
                <input type="submit" value="Upload Single File">
            </form>
        </div>

        <div class="test-section">
            <h2>Multiple Files Upload</h2>
            <p>Upload multiple files at once</p>
            
            <% Response.Write "" %>
            
            <form method="POST" enctype="multipart/form-data">
                <input type="hidden" name="action" value="multiple">
                <p>Select multiple files (allowed: jpg, png, gif, pdf, doc, docx, xls, xlsx):</p>
                <input type="file" name="files" multiple required>
                <input type="submit" value="Upload Multiple Files">
            </form>
        </div>

        <div class="test-section">
            <h2>File Information (Preview)</h2>
            <p>Get file information without uploading</p>
            
            <% Response.Write "" %>
            
            <form method="POST" enctype="multipart/form-data">
                <input type="hidden" name="action" value="info">
                <input type="file" name="infofile" required>
                <input type="submit" value="Get File Info (Without Upload)">
            </form>
        </div>

        <div class="test-section">
            <h2>Features Demonstration</h2>
            
            <h3>Methods Available:</h3>
            <ul>
                <li><strong>Process(fieldName, targetDir, [newFileName])</strong> - Upload single file</li>
                <li><strong>ProcessAll(targetDir)</strong> - Upload all form files</li>
                <li><strong>GetFileInfo(fieldName)</strong> - Get file info without upload</li>
                <li><strong>GetAllFilesInfo()</strong> - Get info for all form files</li>
                <li><strong>BlockExtension(ext)</strong> - Block single extension</li>
                <li><strong>BlockExtensions(exts)</strong> - Block multiple extensions (comma-separated)</li>
                <li><strong>AllowExtension(ext)</strong> - Allow single extension</li>
                <li><strong>AllowExtensions(exts)</strong> - Allow multiple extensions (comma-separated)</li>
                <li><strong>SetUseAllowedOnly(bool)</strong> - Use whitelist mode</li>
            </ul>

            <h3>Properties:</h3>
            <ul>
                <li><strong>MaxFileSize</strong> - Maximum file size in bytes (default: 10MB)</li>
                <li><strong>PreserveOriginalName</strong> - Keep original filename (default: false, use unique name)</li>
                <li><strong>DebugMode</strong> - Enable debug output (default: false)</li>
            </ul>

            <h3>Return Values (Process):</h3>
            <ul>
                <li><strong>IsSuccess</strong> - Boolean indicating success</li>
                <li><strong>OriginalFileName</strong> - Original filename from client</li>
                <li><strong>NewFileName</strong> - Filename on server</li>
                <li><strong>Size</strong> - File size in bytes</li>
                <li><strong>MimeType</strong> - MIME type detected</li>
                <li><strong>Extension</strong> - File extension</li>
                <li><strong>FinalPath</strong> - Absolute path to uploaded file</li>
                <li><strong>RelativePath</strong> - Relative path from www root</li>
                <li><strong>UploadedAt</strong> - Upload timestamp</li>
                <li><strong>ErrorMessage</strong> - Error description if failed</li>
            </ul>
        </div>
    </div>
</body>
</html>
