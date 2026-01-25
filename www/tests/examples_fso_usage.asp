<%
' FileSystemObject Usage Examples
' Complete practical examples using Scripting.FileSystemObject

Response.Write "<h1>FileSystemObject - Practical Examples</h1>"

Set FSO = Server.CreateObject("Scripting.FileSystemObject")

' Example 1: Check if files exist
Response.Write "<h2>Example 1: Check File Existence</h2>"
Response.Write "<pre>"
Response.Write "Set FSO = Server.CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
Response.Write "If FSO.FileExists(""/data/users.txt"") Then" & vbCrLf
Response.Write "    Response.Write ""File found""" & vbCrLf
Response.Write "End If"
Response.Write "</pre>"

' Example 2: List files in a directory
Response.Write "<h2>Example 2: Get File Information</h2>"
Response.Write "<pre>"
Response.Write "Set fileObj = FSO.GetFile(""/test_basics.asp"")" & vbCrLf
Response.Write "Response.Write ""Name: "" & fileObj.Name & vbCrLf" & vbCrLf
Response.Write "Response.Write ""Size: "" & fileObj.Size & "" bytes"" & vbCrLf"
Response.Write "</pre>"
Response.Write "Result:<br>"
If FSO.FileExists("/test_basics.asp") Then
    Set fileObj = FSO.GetFile("/test_basics.asp")
    Response.Write "Name: " & fileObj.Name & "<br>"
    Response.Write "Size: " & fileObj.Size & " bytes<br>"
End If

' Example 3: Path manipulation
Response.Write "<h2>Example 3: Path Manipulation</h2>"
Response.Write "<pre>"
Response.Write "basePath = ""/documents""" & vbCrLf
Response.Write "fileName = ""report.txt""" & vbCrLf
Response.Write "fullPath = FSO.BuildPath(basePath, fileName)"
Response.Write "</pre>"
Response.Write "Result: " & FSO.BuildPath("/documents", "report.txt") & "<br><br>"

' Example 4: Extract filename components
Response.Write "<h2>Example 4: Extract Path Components</h2>"
Response.Write "<pre>"
Response.Write "filePath = ""/www/data/important.doc""" & vbCrLf
Response.Write "Response.Write ""Base Name: "" & FSO.GetBaseName(filePath)" & vbCrLf
Response.Write "Response.Write ""Extension: "" & FSO.GetExtensionName(filePath)" & vbCrLf
Response.Write "Response.Write ""Parent: "" & FSO.GetParentFolderName(filePath)"
Response.Write "</pre>"
Response.Write "Result:<br>"
filePath = "/www/data/important.doc"
Response.Write "Base Name: " & FSO.GetBaseName(filePath) & "<br>"
Response.Write "Extension: " & FSO.GetExtensionName(filePath) & "<br>"
Response.Write "Parent: " & FSO.GetParentFolderName(filePath) & "<br><br>"

' Example 5: Folder operations
Response.Write "<h2>Example 5: Folder Operations</h2>"
Response.Write "<pre>"
Response.Write "If FSO.FolderExists(""/"") Then" & vbCrLf
Response.Write "    Set folderObj = FSO.GetFolder(""/"") " & vbCrLf
Response.Write "    Response.Write ""Folder path: "" & folderObj.Path" & vbCrLf
Response.Write "End If"
Response.Write "</pre>"
Response.Write "Result:<br>"
If FSO.FolderExists("/") Then
    Set folderObj = FSO.GetFolder("/")
    Response.Write "Folder path: " & folderObj.Path & "<br>"
End If

' Example 6: Temporary file name
Response.Write "<h2>Example 6: Generate Temporary Filename</h2>"
Response.Write "<pre>"
Response.Write "tempName = FSO.GetTempName()" & vbCrLf
Response.Write "Response.Write ""Temp file: "" & tempName"
Response.Write "</pre>"
Response.Write "Result: " & FSO.GetTempName() & "<br><br>"

' Example 7: Get absolute path
Response.Write "<h2>Example 7: Get Absolute Path</h2>"
Response.Write "<pre>"
Response.Write "relativePath = ""/data/file.txt""" & vbCrLf
Response.Write "absolutePath = FSO.GetAbsolutePathName(relativePath)" & vbCrLf
Response.Write "Response.Write absolutePath"
Response.Write "</pre>"
Response.Write "Result: " & FSO.GetAbsolutePathName("/data/file.txt") & "<br><br>"

Response.Write "<h2>Common Patterns</h2>"
Response.Write "<h3>Pattern 1: Safe File Reading</h3>"
Response.Write "<pre>"
Response.Write "If FSO.FileExists(""/myfile.txt"") Then" & vbCrLf
Response.Write "    Set textFile = FSO.OpenTextFile(""/myfile.txt"", 1)" & vbCrLf
Response.Write "    content = textFile.ReadAll()" & vbCrLf
Response.Write "    textFile.Close()" & vbCrLf
Response.Write "    Response.Write content" & vbCrLf
Response.Write "End If"
Response.Write "</pre>"

Response.Write "<h3>Pattern 2: File Copy with Overwrite Check</h3>"
Response.Write "<pre>"
Response.Write "sourcePath = ""/source.txt""" & vbCrLf
Response.Write "destPath = ""/backup.txt""" & vbCrLf
Response.Write "If FSO.FileExists(sourcePath) Then" & vbCrLf
Response.Write "    If Not FSO.FileExists(destPath) Then" & vbCrLf
Response.Write "        FSO.CopyFile sourcePath, destPath, False" & vbCrLf
Response.Write "        Response.Write ""File copied successfully""" & vbCrLf
Response.Write "    Else" & vbCrLf
Response.Write "        Response.Write ""Destination file already exists""" & vbCrLf
Response.Write "    End If" & vbCrLf
Response.Write "End If"
Response.Write "</pre>"

Response.Write "<h3>Pattern 3: Create and Write to File</h3>"
Response.Write "<pre>"
Response.Write "Set newFile = FSO.CreateTextFile(""/output.txt"", True)" & vbCrLf
Response.Write "newFile.WriteLine ""Line 1: Hello World""" & vbCrLf
Response.Write "newFile.WriteLine ""Line 2: Welcome to Go-ASP""" & vbCrLf
Response.Write "newFile.Close()" & vbCrLf
Response.Write "Response.Write ""File created successfully"""
Response.Write "</pre>"

Response.Write "<h2>Supported CreateObject Aliases</h2>"
Response.Write "<pre>"
Response.Write "' Classic ASP syntax (recommended)" & vbCrLf
Response.Write "Set FSO = Server.CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "' Direct G3Files library (faster)" & vbCrLf
Response.Write "Set FSO = Server.CreateObject(""G3FILES"")"
Response.Write "</pre>"

Response.Write "<h2>All Available Methods</h2>"
Response.Write "<ul>"
Response.Write "<li>FileExists(path) - Boolean</li>"
Response.Write "<li>GetFile(path) - FSOFile object</li>"
Response.Write "<li>DeleteFile(path) - Void</li>"
Response.Write "<li>CopyFile(src, dst, overwrite) - Void</li>"
Response.Write "<li>MoveFile(src, dst) - Void</li>"
Response.Write "<li>CreateTextFile(filename, overwrite, unicode) - TextStream</li>"
Response.Write "<li>OpenTextFile(filename, iomode, create, format) - TextStream</li>"
Response.Write "<li>FolderExists(path) - Boolean</li>"
Response.Write "<li>GetFolder(path) - FSOFolder object</li>"
Response.Write "<li>CreateFolder(path) - FSOFolder object</li>"
Response.Write "<li>DeleteFolder(path) - Void</li>"
Response.Write "<li>MoveFolder(src, dst) - Void</li>"
Response.Write "<li>BuildPath(base, rel) - String</li>"
Response.Write "<li>GetFileName(path) - String</li>"
Response.Write "<li>GetBaseName(filename) - String</li>"
Response.Write "<li>GetExtensionName(filename) - String</li>"
Response.Write "<li>GetParentFolderName(path) - String</li>"
Response.Write "<li>GetAbsolutePathName(path) - String</li>"
Response.Write "<li>GetDriveName(path) - String</li>"
Response.Write "<li>GetTempName() - String</li>"
Response.Write "</ul>"

Response.Write "<h2>Summary</h2>"
Response.Write "<p>The FileSystemObject provides complete file and folder manipulation capabilities</p>"
Response.Write "<p>All paths are relative to the web application root for security</p>"
Response.Write "<p>Use Server.MapPath() to convert virtual paths if needed</p>"
%>
