<!--#include file="flexforms.asp" -->

<%
' 1. Instatiate the FlexForms class
Dim form
Set form = New FlexForms

' Crucial security settings
form.SetSecretKey "YourSuperSecureSecretKey123!"
form.SetTokenExtraInfo "user_session_or_ip"

' 2. Initialize the required Options and Errors dictionaries for the engine
Dim options, errors
Set options = CreateObject("Scripting.Dictionary")
Set errors = CreateObject("Scripting.Dictionary")

' Form structural settings
options.Add "useform", True
options.Add "formmode", "post"
options.Add "submit", "Save Information" ' Submit button text

' 3. Build the form fields using Scripting.Dictionary
Dim fieldName, fieldEmail, fieldPassword, fieldBio, fieldFile, fieldTable
Dim arrayCampos(5) ' Array of fields

' Field: Full Name
Set fieldName = CreateObject("Scripting.Dictionary")
fieldName.Add "type", "text"
fieldName.Add "name", "txt_name"
fieldName.Add "title", "Full Name"
fieldName.Add "value", ""
Set arrayCampos(0) = fieldName

' Field: Email
Set fieldEmail = CreateObject("Scripting.Dictionary")
fieldEmail.Add "type", "text"
fieldEmail.Add "name", "txt_email"
fieldEmail.Add "title", "Professional Email"
fieldEmail.Add "value", ""
Set arrayCampos(1) = fieldEmail

' Field: Bio
Set fieldBio = CreateObject("Scripting.Dictionary")
fieldBio.Add "type", "textarea"
fieldBio.Add "name", "txt_bio"
fieldBio.Add "title", "Short Bio"
fieldBio.Add "value", ""
Set arrayCampos(2) = fieldBio

' Field: Attachment
Set fieldFile = CreateObject("Scripting.Dictionary")
fieldFile.Add "type", "file"
fieldFile.Add "name", "file_upload"
fieldFile.Add "title", "Profile Picture"
Set arrayCampos(3) = fieldFile

' Field: Password
Set fieldPassword = CreateObject("Scripting.Dictionary")
fieldPassword.Add "type", "password"
fieldPassword.Add "name", "txt_password"
fieldPassword.Add "title", "Access Password"
fieldPassword.Add "value", ""
Set arrayCampos(4) = fieldPassword

' Field: Table
Set fieldTable = CreateObject("Scripting.Dictionary")
fieldTable.Add "type", "table"
fieldTable.Add "name", "tbl_data"
fieldTable.Add "title", "User List"
fieldTable.Add "cols", Array("ID", "Name", "Role")
fieldTable.Add "rows", Array(Array(1, "John Doe", "Admin"), Array(2, "Jane Smith", "User"))
Set arrayCampos(5) = fieldTable

'  Bind the collection of fields to the global options dictionary
options.Add "fields", arrayCampos

' 4. Process POST logic and error validations
Dim successMsg
successMsg = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Use G3FILEUPLOADER to handle multipart/form-data
    Dim uploader, fileResult, valName, valEmail, valBio, fs
    Set uploader = Server.CreateObject("G3FILEUPLOADER")
    Set fs = Server.CreateObject("G3FILES")
    
    ' Ensure temp directory exists
    Call fs.Mkdir(Server.MapPath("temp"))
    
    ' Get text field values from uploader (since it's a multipart request)
    valName = uploader.Form("txt_name")
    valEmail = uploader.Form("txt_email")
    valBio = uploader.Form("txt_bio")
    
    ' Update field values so they are preserved in the form
    fieldName("value") = valName
    fieldEmail("value") = valEmail
    fieldBio("value") = valBio

    ' Simple validations
    If valName = "" Then errors.Add "txt_name", "Please provide your full name."
    If valEmail = "" Then errors.Add "txt_email", "Please provide a valid email address."
    
    If errors.Count = 0 Then
        ' Check if a file was actually provided before processing
        Dim fileInfo
        Set fileInfo = uploader.GetFileInfo("file_upload")
        
        ' G3FILEUPLOADER returns IsSuccess only on failure or ErrorMessage. 
        ' On success, it returns OriginalFileName.
        If fileInfo.Exists("OriginalFileName") Then
            ' Process file upload - Pass "temp" as virtual path
            Set fileResult = uploader.Process("file_upload", "temp")
            
            If fileResult("IsSuccess") Then
                successMsg = "Form processed! User: " & valName & ". File saved as: " & fileResult("NewFileName")
            Else
                errors.Add "file_upload", "Upload error: " & fileResult("ErrorMessage")
            End If
        Else
            ' No file uploaded, but we can still succeed for the rest of the form
            successMsg = "Form processed! User: " & valName & " (No file was selected)."
        End If
    End If
End If
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>FormTest — AxonASP</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; margin: 50px; }
        .container { max-width: 600px; margin: 0 auto; background: #ffffff; padding: 30px; border-radius: 6px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        h2 { margin-top: 0; color: #2c3e50; border-bottom: 2px solid #ecf0f1; padding-bottom: 10px; }
        
        /* Classes CSS geradas dinamicamente pelos seletores do FlexForms */
        .formitem { margin-bottom: 18px; }
        .formitemtitle { font-weight: 600; margin-bottom: 6px; color: #34495e; font-size: 14px; }
        .text, .textarea { width: 100%; padding: 10px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; transition: border 0.2s; }
        .text:focus, .textarea:focus { border-color: #3498db; outline: none; }
        .textarea { height: 100px; resize: vertical; }
        .submit { background-color: #2ecc71; color: white; border: none; padding: 12px 20px; border-radius: 4px; font-size: 15px; cursor: pointer; font-weight: bold; width: 100%; }
        .submit:hover { background-color: #27ae60; }
        .formitemerror { color: #e74c3c; font-size: 12px; margin-top: 4px; }
        
        /* Mensagens de feedback globais do FlexForms */
        .ff_formmessagewrap { margin-bottom: 20px; }
        .message { padding: 12px; border-radius: 4px; font-size: 14px; font-weight: 500; }
        .messageerror { background-color: #fde8e8; color: #e74c3c; border-left: 4px solid #e74c3c; }
        .messagesuccess { background-color: #edfbf7; color: #2ecc71; border-left: 4px solid #2ecc71; }
        
        /* Table styles */
        .formitemtable { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 13px; }
        .formitemtable th, .formitemtable td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        .formitemtable th { background-color: #f8f9fa; font-weight: bold; }
        .row.altrow { background-color: #fcfcfc; }
    </style>
</head>
<body>

<div class="container">
    <h2>User Profile Registration</h2>
    
    <%
    ' 5. Monitoring and rendering global system notifications
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
        If errors.Count = 0 Then
            form.OutputMessage "success", successMsg
        Else
            form.OutputMessage "error", "There were problems processing your form."
        End If
    End If

    ' 6. Invocation of the converted class rendering engine
    form.Generate options, errors, True
    %>

</div>

</body>
</html>
<%
' 7. Explicit memory deallocation
Set fieldName = Nothing
Set fieldEmail = Nothing
Set fieldBio = Nothing
Set fieldFile = Nothing
Set fieldPassword = Nothing
Set fieldTable = Nothing
Set options = Nothing
Set errors = Nothing
Set form = Nothing
%>