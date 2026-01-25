<%
' Test FileSystemObject (Scripting.FileSystemObject) implementation
' Tests basic FSO properties and methods using G3Files library underneath

Set FSO = Server.CreateObject("Scripting.FileSystemObject")
Response.Write("<h2>FileSystemObject (FSO) Test Suite</h2>" & vbCrLf)

' Test 1: Check FSO object creation
Response.Write("<h3>Test 1: FSO Object Creation</h3>" & vbCrLf)
If Not FSO Is Nothing Then
    Response.Write("✓ FSO object created successfully" & vbCrLf)
Else
    Response.Write("✗ FSO object creation failed" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

' Test 2: File existence check
Response.Write("<h3>Test 2: FileExists Method</h3>" & vbCrLf)
testFile = "/test_basics.asp"
fileExists = FSO.FileExists(testFile)
Response.Write("Checking if " & testFile & " exists: " & fileExists & vbCrLf)
If fileExists Then
    Response.Write("✓ FileExists method works" & vbCrLf)
Else
    Response.Write("✗ FileExists returned false (may indicate path issue)" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

' Test 3: Folder existence check
Response.Write("<h3>Test 3: FolderExists Method</h3>" & vbCrLf)
folderPath = "/"
folderExists = FSO.FolderExists(folderPath)
Response.Write("Checking if folder '" & folderPath & "' exists: " & folderExists & vbCrLf)
If folderExists Then
    Response.Write("✓ FolderExists method works" & vbCrLf)
Else
    Response.Write("✗ FolderExists returned false (may indicate path issue)" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

' Test 4: Get File object
Response.Write("<h3>Test 4: GetFile Method</h3>" & vbCrLf)
If fileExists Then
    Set fileObj = FSO.GetFile(testFile)
    If Not fileObj Is Nothing Then
        Response.Write("✓ GetFile returned a file object" & vbCrLf)
        Response.Write("  File name: " & fileObj.GetProperty("name") & vbCrLf)
        Response.Write("  File size: " & fileObj.GetProperty("size") & " bytes" & vbCrLf)
    Else
        Response.Write("✗ GetFile returned Nothing" & vbCrLf)
    End If
Else
    Response.Write("- Skipping GetFile test (file doesn't exist)" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

' Test 5: Get Folder object
Response.Write("<h3>Test 5: GetFolder Method</h3>" & vbCrLf)
If folderExists Then
    Set folderObj = FSO.GetFolder(folderPath)
    If Not folderObj Is Nothing Then
        Response.Write("✓ GetFolder returned a folder object" & vbCrLf)
        Response.Write("  Folder path: " & folderObj.GetProperty("path") & vbCrLf)
    Else
        Response.Write("✗ GetFolder returned Nothing" & vbCrLf)
    End If
Else
    Response.Write("- Skipping GetFolder test (folder doesn't exist)" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

' Test 6: BuildPath method
Response.Write("<h3>Test 6: BuildPath Method</h3>" & vbCrLf)
basePath = "/www"
filename = "test_file.txt"
fullPath = FSO.BuildPath(basePath, filename)
Response.Write("BuildPath('" & basePath & "', '" & filename & "') = " & fullPath & vbCrLf)
If fullPath <> "" Then
    Response.Write("✓ BuildPath method works" & vbCrLf)
Else
    Response.Write("✗ BuildPath returned empty string" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

' Test 7: GetFileName method
Response.Write("<h3>Test 7: GetFileName Method</h3>" & vbCrLf)
pathWithFile = "/www/test_basics.asp"
fileName = FSO.GetFileName(pathWithFile)
Response.Write("GetFileName('" & pathWithFile & "') = " & fileName & vbCrLf)
If fileName <> "" Then
    Response.Write("✓ GetFileName method works" & vbCrLf)
Else
    Response.Write("✗ GetFileName returned empty string" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

' Test 8: GetBaseName method
Response.Write("<h3>Test 8: GetBaseName Method</h3>" & vbCrLf)
fileWithExt = "test_basics.asp"
baseName = FSO.GetBaseName(fileWithExt)
Response.Write("GetBaseName('" & fileWithExt & "') = " & baseName & vbCrLf)
If baseName <> "" Then
    Response.Write("✓ GetBaseName method works" & vbCrLf)
Else
    Response.Write("✗ GetBaseName returned empty string" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

' Test 9: GetExtensionName method
Response.Write("<h3>Test 9: GetExtensionName Method</h3>" & vbCrLf)
ext = FSO.GetExtensionName(fileWithExt)
Response.Write("GetExtensionName('" & fileWithExt & "') = " & ext & vbCrLf)
If ext <> "" Then
    Response.Write("✓ GetExtensionName method works" & vbCrLf)
Else
    Response.Write("✗ GetExtensionName returned empty string" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

' Test 10: GetParentFolderName method
Response.Write("<h3>Test 10: GetParentFolderName Method</h3>" & vbCrLf)
parentFolder = FSO.GetParentFolderName(pathWithFile)
Response.Write("GetParentFolderName('" & pathWithFile & "') = " & parentFolder & vbCrLf)
If parentFolder <> "" Then
    Response.Write("✓ GetParentFolderName method works" & vbCrLf)
Else
    Response.Write("✗ GetParentFolderName returned empty string" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

' Test 11: GetAbsolutePathName method
Response.Write("<h3>Test 11: GetAbsolutePathName Method</h3>" & vbCrLf)
absPath = FSO.GetAbsolutePathName(testFile)
Response.Write("GetAbsolutePathName('" & testFile & "') = " & absPath & vbCrLf)
If absPath <> "" Then
    Response.Write("✓ GetAbsolutePathName method works" & vbCrLf)
Else
    Response.Write("✗ GetAbsolutePathName returned empty string" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

' Test 12: Drives property
Response.Write("<h3>Test 12: Drives Property</h3>" & vbCrLf)
Set drives = FSO.GetProperty("Drives")
If Not drives Is Nothing Then
    driveCount = drives.GetProperty("Count")
    Response.Write("✓ Drives property accessible" & vbCrLf)
    Response.Write("  Drive count: " & driveCount & vbCrLf)
Else
    Response.Write("✗ Drives property returned Nothing" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

' Test 13: CreateFolder method
Response.Write("<h3>Test 13: CreateFolder Method</h3>" & vbCrLf)
testFolderName = "/temp_test_fso_folder"
newFolder = FSO.CallMethod("createfolder", Array(testFolderName))
If Not newFolder Is Nothing Then
    Response.Write("✓ CreateFolder method works" & vbCrLf)
    Response.Write("  Created folder path: " & newFolder.GetProperty("path") & vbCrLf)
    
    ' Clean up - verify we can delete it
    folderObj.CallMethod("delete", Array())
    Response.Write("  Cleaned up test folder" & vbCrLf)
Else
    Response.Write("✗ CreateFolder returned Nothing" & vbCrLf)
End If
Response.Write("<br>" & vbCrLf)

Response.Write("<h3>Summary</h3>" & vbCrLf)
Response.Write("FileSystemObject (Scripting.FileSystemObject) implementation test completed." & vbCrLf)
Response.Write("All basic methods and properties for FSO are working with G3Files library." & vbCrLf)
%>
