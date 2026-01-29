<%
Option Explicit

Dim isPost, action, uploader, result, uploadDir, results, i, fileInfo
Dim simpleSuccess, simpleError, multipleSuccess, multipleError, infoSuccess, infoError
Dim simpleResultObj, multipleResultsArr, fileInfoObj
Dim statusText, statusClass, resultHtml

isPost = (Request.ServerVariables("REQUEST_METHOD") = "POST")
action = Request.Form("action")
resultHtml = ""

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
    ' ENABLE DEBUG MODE
    uploader.SetProperty "debugmode", True
    uploader.BlockExtensions "exe,dll,bat,cmd"
    uploader.SetProperty "maxfilesize", 50485760  ' 50MB
    
    Set simpleResultObj = uploader.Process("file1", "/uploads")
    
    If simpleResultObj("IsSuccess") Then
        simpleSuccess = True
        resultHtml = "<div class='result success'>" & _
                    "<strong>Upload Successful!</strong><br>" & _
                    "Original Name: " & simpleResultObj("OriginalFileName") & "<br>" & _
                    "New Name: " & simpleResultObj("NewFileName") & "<br>" & _
                    "Size: " & simpleResultObj("Size") & " bytes<br>" & _
                    "MIME Type: " & simpleResultObj("MimeType") & "<br>" & _
                    "Extension: " & simpleResultObj("Extension") & "<br>" & _
                    "Relative Path: " & simpleResultObj("RelativePath") & "</div>"
    Else
        simpleSuccess = False
        simpleError = simpleResultObj("ErrorMessage")
        resultHtml = "<div class='result error'>" & _
                    "<strong>Upload Failed!</strong><br>" & _
                    "Error: " & simpleError & "</div>"
    End If
End If

' MULTIPLE FILES UPLOAD HANDLING
If isPost And action = "multiple" Then
    Set uploader = Server.CreateObject("G3FileUploader")
    uploader.SetProperty "debugmode", True
    uploader.AllowExtensions "jpg,png,gif,pdf,doc,docx,xls,xlsx,txt"
    uploader.SetUseAllowedOnly(True)
    
    Set multipleResultsArr = uploader.ProcessAll("/uploads")
    
    If IsArray(multipleResultsArr) And UBound(multipleResultsArr) >= 0 Then
        multipleSuccess = True
        resultHtml = "<table><tr><th>File Name</th><th>Size (bytes)</th><th>MIME Type</th><th>Extension</th><th>Status</th></tr>"
        For i = 0 To UBound(multipleResultsArr)
            If multipleResultsArr(i)("IsSuccess") Then
                statusText = "OK"
                statusClass = "success"
            Else
                statusText = "FAILED: " & multipleResultsArr(i)("ErrorMessage")
                statusClass = "error"
            End If
            resultHtml = resultHtml & "<tr><td>" & multipleResultsArr(i)("OriginalFileName") & "</td>" & _
                        "<td>" & multipleResultsArr(i)("Size") & "</td>" & _
                        "<td>" & multipleResultsArr(i)("MimeType") & "</td>" & _
                        "<td>" & multipleResultsArr(i)("Extension") & "</td>" & _
                        "<td class='" & statusClass & "'>" & statusText & "</td></tr>"
        Next
        resultHtml = resultHtml & "</table>"
    Else
        multipleSuccess = False
        multipleError = "No files uploaded"
        resultHtml = "<div class='result error'>" & multipleError & "</div>"
    End If
End If

' FILE INFO HANDLING
If isPost And action = "info" Then
    Set uploader = Server.CreateObject("G3FileUploader")
    uploader.SetProperty "debugmode", True
    uploader.BlockExtensions "exe,dll,bat,cmd,vbs,js,scr"
    
    Set fileInfoObj = uploader.GetFileInfo("infofile")
    
    If fileInfoObj("IsSuccess") = False Then
        infoSuccess = False
        infoError = fileInfoObj("ErrorMessage")
        resultHtml = "<div class='result error'>" & infoError & "</div>"
    Else
        infoSuccess = True
        resultHtml = "<div class='result info'>" & _
                    "<strong>File Information:</strong><br>" & _
                    "File Name: " & fileInfoObj("OriginalFileName") & "<br>" & _
                    "Size: " & fileInfoObj("Size") & " bytes<br>" & _
                    "MIME Type: " & fileInfoObj("MimeType") & "<br>" & _
                    "Extension: " & fileInfoObj("Extension") & "<br>"
        If fileInfoObj("IsValid") Then
            resultHtml = resultHtml & "Is Valid: Yes<br>"
        Else
            resultHtml = resultHtml & "Is Valid: No<br>"
        End If
        If fileInfoObj("ExceedsMaxSize") Then
            resultHtml = resultHtml & "Exceeds Max Size: Yes"
        Else
            resultHtml = resultHtml & "Exceeds Max Size: No"
        End If
        resultHtml = resultHtml & "</div>"
    End If
End If
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>G3FileUploader - Test (DEBUG MODE)</title>
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
        .debug-info { background-color: #ffffcc; padding: 10px; margin-top: 20px; border: 1px solid #cccc00; border-radius: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3FileUploader - Upload Test (DEBUG MODE ENABLED)</h1>
        
        <div class="debug-info">
            <strong>DEBUG MODE ENABLED:</strong> The server console will show detailed debug information for each upload attempt.
            Check the console logs to see ParseMultipartForm events and FormFile calls.
        </div>
        
        <div class="test-section">
            <h2>Simple File Upload</h2>
            <p>Upload a single file to the /uploads directory</p>
            <p style="color: #666; font-size: 0.9em;">Max file size: 50MB | Blocked: exe, dll, bat, cmd</p>
            
            <% If action = "simple" Then Response.Write resultHtml End If %>
            
            <form method="POST" enctype="multipart/form-data">
                <input type="hidden" name="action" value="simple">
                <input type="file" name="file1" required>
                <input type="submit" value="Upload Single File">
            </form>
        </div>

        <div class="test-section">
            <h2>Multiple Files Upload</h2>
            <p>Upload multiple files at once</p>
            <p style="color: #666; font-size: 0.9em;">Allowed: jpg, png, gif, pdf, doc, docx, xls, xlsx</p>
            
            <% If action = "multiple" Then Response.Write resultHtml End If %>
            
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
            
            <% If action = "info" Then Response.Write resultHtml End If %>
            
            <form method="POST" enctype="multipart/form-data">
                <input type="hidden" name="action" value="info">
                <input type="file" name="infofile" required>
                <input type="submit" value="Get File Info (Without Upload)">
            </form>
        </div>

        <div class="test-section">
            <h2>Troubleshooting</h2>
            
            <h3>If files are not uploading:</h3>
            <ol>
                <li><strong>Check the console logs:</strong> Look for [G3FileUploader DEBUG] messages</li>
                <li><strong>Verify field names:</strong> Make sure the HTML input field names match your code (file1, files, infofile)</li>
                <li><strong>Check permissions:</strong> Ensure /uploads and /temp/uploads directories exist and are writable</li>
                <li><strong>Monitor file size:</strong> Large files may take time to upload</li>
            </ol>

            <h3>Common Issues:</h3>
            <ul>
                <li><strong>Small files not received:</strong> Check console for "File field not found" messages</li>
                <li><strong>Large files timeout:</strong> Increase script timeout in server config</li>
                <li><strong>Connection reset:</strong> May indicate a server-side error - check console logs</li>
            </ul>

            <h3>Debug Output Example:</h3>
            <pre style="background-color: #f5f5f5; padding: 10px; border-radius: 3px;">
[G3FileUploader DEBUG] Available file fields: file1
[G3FileUploader DEBUG] FormFile called for: file1
[G3FileUploader DEBUG] io.Copy result: 1024 bytes written
            </pre>
        </div>
    </div>
</body>
</html>
