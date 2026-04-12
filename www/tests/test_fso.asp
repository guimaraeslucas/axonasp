<%
@ Language = VBScript
%>
<%
Option Explicit

Dim fso
Dim passCount
Dim failCount

Dim testsPhysical
Dim rootPhysical
Dim basePhysical
Dim nestedPhysical
Dim nestedRenamedPhysical
Dim copyPhysical
Dim movedPhysical
Dim alphaPhysical
Dim betaPhysical
Dim betaRenamedPhysical
Dim gammaPhysical
Dim linesPhysical
Dim streamPhysical
Dim folderCreatedPhysical
Dim deletePhysical

Set fso = Server.CreateObject("Scripting.FileSystemObject")
passCount = 0
failCount = 0

testsPhysical = Server.MapPath(".")
rootPhysical = Server.MapPath("/")
basePhysical = Server.MapPath("fso_runtime")
nestedPhysical = Server.MapPath("fso_runtime\\nested")
nestedRenamedPhysical = Server.MapPath("fso_runtime\\nested_renamed")
copyPhysical = Server.MapPath("fso_runtime_copy")
movedPhysical = Server.MapPath("fso_runtime_moved")
alphaPhysical = Server.MapPath("fso_runtime\\alpha.txt")
betaPhysical = Server.MapPath("fso_runtime\\beta.txt")
betaRenamedPhysical = Server.MapPath("fso_runtime\\beta_renamed.txt")
gammaPhysical = Server.MapPath("fso_runtime\\gamma.txt")
linesPhysical = Server.MapPath("fso_runtime\\lines.txt")
streamPhysical = Server.MapPath("fso_runtime\\stream_ops.txt")
folderCreatedPhysical = Server.MapPath("fso_runtime\\folder_created.txt")
deletePhysical = Server.MapPath("fso_runtime\\delete_me.txt")

Sub WriteLine(text)
    Response.Write text & vbCrLf
End Sub

Sub LogResult(testName, ok, details)
    If ok Then
        passCount = passCount + 1
        WriteLine "PASS | " & testName & " | " & details
    Else
        failCount = failCount + 1
        WriteLine "FAIL | " & testName & " | " & details
    End If
End Sub

Function NormalizePath(value)
    NormalizePath = LCase(Replace(CStr(value), "/", "\\"))
End Function

Function SamePath(leftValue, rightValue)
    SamePath = (NormalizePath(leftValue) = NormalizePath(rightValue))
End Function

Sub AssertTrue(testName, condition, details)
    LogResult testName, CBool(condition), details
End Sub

Sub AssertEqual(testName, expected, actual)
    Dim ok
    ok = (CStr(expected) = CStr(actual))
    LogResult testName, ok, "expected=[" & CStr(expected) & "] actual=[" & CStr(actual) & "]"
End Sub

Sub AssertPathEqual(testName, expected, actual)
    LogResult testName, SamePath(expected, actual), "expected=[" & CStr(expected) & "] actual=[" & CStr(actual) & "]"
End Sub

Sub CleanupSandbox()
    If fso.FolderExists(movedPhysical) Then
        fso.DeleteFolder movedPhysical
    End If
    If fso.FolderExists(copyPhysical) Then
        fso.DeleteFolder copyPhysical
    End If
    If fso.FolderExists(basePhysical) Then
        fso.DeleteFolder basePhysical
    End If
End Sub

Sub CreateAlphaFile()
    Dim stream
    Set stream = fso.CreateTextFile(alphaPhysical, True)
    stream.Write "Alpha payload"
    stream.Close
    Set stream = Nothing
End Sub

Sub CreateLinesFile()
    Dim stream
    Set stream = fso.CreateTextFile(linesPhysical, True)
    stream.WriteLine "first"
    stream.WriteLine "second"
    stream.WriteBlankLines 1
    stream.Write "third"
    stream.Close
    Set stream = Nothing
End Sub

Sub CreateDeleteFile()
    Dim stream
    Set stream = fso.CreateTextFile(deletePhysical, True)
    stream.Write "delete"
    stream.Close
    Set stream = Nothing
End Sub

Function ReadAllText(path)
    Dim stream
    Set stream = fso.OpenTextFile(path, 1)
    ReadAllText = stream.ReadAll
    stream.Close
    Set stream = Nothing
End Function

If fso.FolderExists(movedPhysical) Then
    fso.DeleteFolder movedPhysical
End If
If fso.FolderExists(copyPhysical) Then
    fso.DeleteFolder copyPhysical
End If
If fso.FolderExists(basePhysical) Then
    fso.DeleteFolder basePhysical
End If

Dim baseFolder
Dim nestedFolder
Dim copiedFolder
Dim movedFolder
Dim alphaFile
Dim betaFile
Dim gammaFile
Dim DeleteFile
Dim folderTextStream
Dim writeStream
Dim readLineStream
Dim readAllStream
Dim readChunkStream
Dim fileStream
Dim drives
Dim driveFromMethod
Dim driveFromCollection
Dim alphaDrive
Dim alphaParentFolder
Dim baseDrive
Dim baseParentFolder
Dim alphaFromCollection
Dim nestedFromCollection
Dim driveName
Dim driveLetter
Dim RootFolder
Dim tempName
Dim builtPath
Dim allText
Dim lineValue
Dim chunkValue
Dim folderFileCount
Dim folderSubFolderCount

Response.ContentType = "text/plain"
WriteLine "AxonASP Classic ASP FileSystemObject extensive test"
WriteLine String(72, "=")

WriteLine "SECTION | Root methods and path helpers"
Set baseFolder = fso.CreateFolder(basePhysical)
AssertTrue "CreateFolder creates base folder", fso.FolderExists(basePhysical), basePhysical
AssertPathEqual "CreateFolder returns folder object path", basePhysical, baseFolder.Path

builtPath = fso.BuildPath(basePhysical, "joined.txt")
AssertPathEqual "BuildPath joins folder and file", Server.MapPath("fso_runtime\\joined.txt"), builtPath
AssertEqual "GetFileName extracts file name", "alpha.txt", fso.GetFileName(alphaPhysical)
AssertEqual "GetBaseName extracts base name", "alpha", fso.GetBaseName(alphaPhysical)
AssertEqual "GetExtensionName extracts extension", "txt", fso.GetExtensionName(alphaPhysical)
AssertPathEqual "GetParentFolderName resolves parent", basePhysical, fso.GetParentFolderName(alphaPhysical)
AssertPathEqual "GetAbsolutePathName resolves relative path", alphaPhysical, fso.GetAbsolutePathName("fso_runtime\\alpha.txt")

tempName = fso.GetTempName()
AssertTrue "GetTempName returns non-empty value", (Len(tempName) > 0), tempName
AssertTrue "GetSpecialFolder(0) returns non-empty value", (Len(CStr(fso.GetSpecialFolder(0))) > 0), CStr(fso.GetSpecialFolder(0))
AssertTrue "GetSpecialFolder(1) returns non-empty value", (Len(CStr(fso.GetSpecialFolder(1))) > 0), CStr(fso.GetSpecialFolder(1))
AssertTrue "GetSpecialFolder(2) returns non-empty value", (Len(CStr(fso.GetSpecialFolder(2))) > 0), CStr(fso.GetSpecialFolder(2))

driveName = fso.GetDriveName(alphaPhysical)
AssertTrue "GetDriveName returns non-empty drive", (Len(driveName) > 0), driveName
AssertTrue "DriveExists returns true for current drive", fso.DriveExists(driveName), driveName

Set driveFromMethod = fso.GetDrive(driveName)
AssertEqual "Drive.DriveLetter matches current drive", UCase(Left(driveName, 1)), UCase(driveFromMethod.DriveLetter)
AssertTrue "Drive.Path is non-empty", (Len(CStr(driveFromMethod.Path)) > 0), CStr(driveFromMethod.Path)
AssertTrue "Drive.IsReady is true", CBool(driveFromMethod.IsReady), CStr(driveFromMethod.IsReady)
AssertTrue "Drive.FileSystem is non-empty", (Len(CStr(driveFromMethod.FileSystem)) > 0), CStr(driveFromMethod.FileSystem)
AssertTrue "Drive.SerialNumber is non-empty", (Len(CStr(driveFromMethod.SerialNumber)) > 0), CStr(driveFromMethod.SerialNumber)
AssertTrue "Drive.ShareName is non-empty", (Len(CStr(driveFromMethod.ShareName)) > 0), CStr(driveFromMethod.ShareName)
AssertTrue "Drive.VolumeName is non-empty", (Len(CStr(driveFromMethod.VolumeName)) > 0), CStr(driveFromMethod.VolumeName)
AssertTrue "Drive.TotalSize is populated", (Len(CStr(driveFromMethod.TotalSize)) > 0), CStr(driveFromMethod.TotalSize)
AssertTrue "Drive.FreeSpace is populated", (Len(CStr(driveFromMethod.FreeSpace)) > 0), CStr(driveFromMethod.FreeSpace)
AssertTrue "Drive.AvailableSpace is populated", (Len(CStr(driveFromMethod.AvailableSpace)) > 0), CStr(driveFromMethod.AvailableSpace)

Set RootFolder = driveFromMethod.RootFolder
AssertTrue "Drive.RootFolder reports root folder", CBool(RootFolder.IsRootFolder), CStr(RootFolder.Path)

Set drives = fso.Drives
AssertTrue "Drives.Count is positive", (CLng(drives.Count) > 0), CStr(drives.Count)
Set driveFromCollection = drives.Item(driveName)
AssertEqual "Drives.Item returns same drive letter", UCase(driveFromMethod.DriveLetter), UCase(driveFromCollection.DriveLetter)

WriteLine ""
WriteLine "SECTION | TextStream create, write, read, and property access"
Set writeStream = fso.CreateTextFile(streamPhysical, True)
AssertEqual "CreateTextFile initial line", 1, writeStream.Line
AssertEqual "CreateTextFile initial column", 1, writeStream.Column
writeStream.Write "AB"
AssertEqual "Write updates column", 3, writeStream.Column
writeStream.WriteLine "CD"
AssertEqual "WriteLine advances line", 2, writeStream.Line
AssertEqual "WriteLine resets column", 1, writeStream.Column
writeStream.WriteBlankLines 2
AssertEqual "WriteBlankLines advances line count", 4, writeStream.Line
writeStream.Write "EF"
AssertEqual "Write after blank lines updates column", 3, writeStream.Column
writeStream.Close

AssertTrue "FileExists after CreateTextFile", fso.FileExists(streamPhysical), streamPhysical

Set readLineStream = fso.OpenTextFile(streamPhysical, 1)
AssertEqual "OpenTextFile read stream initial line", 1, readLineStream.Line
AssertEqual "OpenTextFile read stream initial column", 1, readLineStream.Column
lineValue = readLineStream.ReadLine
AssertEqual "ReadLine method returns first line", "ABCD", lineValue
AssertEqual "ReadLine increments line property", 2, readLineStream.Line
lineValue = readLineStream.ReadLine
AssertEqual "ReadLine property returns blank line", "", lineValue
lineValue = readLineStream.ReadLine
AssertEqual "ReadLine property returns second blank line", "", lineValue
lineValue = readLineStream.ReadLine
AssertEqual "ReadLine property returns final line", "EF", lineValue
AssertTrue "AtEndOfStream becomes true after final line", CBool(readLineStream.AtEndOfStream), CStr(readLineStream.AtEndOfStream)
readLineStream.Close

Set readAllStream = fso.OpenTextFile(streamPhysical, 1)
allText = readAllStream.ReadAll
AssertEqual "ReadAll property returns complete content", "ABCD" & vbCrLf & vbCrLf & vbCrLf & "EF", allText
AssertTrue "ReadAll marks stream end", CBool(readAllStream.AtEndOfStream), CStr(readAllStream.AtEndOfStream)
readAllStream.Close

Set readChunkStream = fso.OpenTextFile(streamPhysical, 1)
chunkValue = readChunkStream.Read(4)
AssertEqual "Read returns requested characters", "ABCD", chunkValue
AssertEqual "Read updates column", 5, readChunkStream.Column
readChunkStream.Skip 6
AssertEqual "Skip advances column", 11, readChunkStream.Column
chunkValue = readChunkStream.Read(2)
AssertEqual "Read after Skip continues from correct position", "EF", chunkValue
readChunkStream.Close

Set alphaFile = fso.GetFile(streamPhysical)

WriteLine ""
WriteLine "SECTION | File object properties and methods"
AssertTrue "GetFile returns existing file object", fso.FileExists(streamPhysical), streamPhysical
AssertEqual "File.Name", "stream_ops.txt", alphaFile.Name
AssertTrue "File.ShortName is non-empty", (Len(CStr(alphaFile.ShortName)) > 0), CStr(alphaFile.ShortName)
AssertPathEqual "File.Path", streamPhysical, alphaFile.Path
AssertTrue "File.ShortPath is non-empty", (Len(CStr(alphaFile.ShortPath)) > 0), CStr(alphaFile.ShortPath)
AssertTrue "File.Attributes is non-negative", (CLng(alphaFile.Attributes) >= 0), CStr(alphaFile.Attributes)
AssertTrue "File.DateCreated is non-empty", (Len(CStr(alphaFile.DateCreated)) > 0), CStr(alphaFile.DateCreated)
AssertTrue "File.DateLastAccessed is non-empty", (Len(CStr(alphaFile.DateLastAccessed)) > 0), CStr(alphaFile.DateLastAccessed)
AssertTrue "File.DateLastModified is non-empty", (Len(CStr(alphaFile.DateLastModified)) > 0), CStr(alphaFile.DateLastModified)
Set alphaDrive = alphaFile.Drive
AssertEqual "File.Drive.DriveLetter", UCase(driveFromMethod.DriveLetter), UCase(alphaDrive.DriveLetter)
AssertTrue "File.Size is positive", (CLng(alphaFile.Size) > 0), CStr(alphaFile.Size)
AssertEqual "File.Type", "TXT File", alphaFile.Type
AssertTrue "File.IsRootFolder is false", (Not CBool(alphaFile.IsRootFolder)), CStr(alphaFile.IsRootFolder)
Set alphaParentFolder = alphaFile.ParentFolder
AssertPathEqual "File.ParentFolder.Path", basePhysical, alphaParentFolder.Path
AssertTrue "GetFileVersion returns non-empty string", (Len(CStr(fso.GetFileVersion(streamPhysical))) > 0), CStr(fso.GetFileVersion(streamPhysical))

alphaFile.Copy betaPhysical, True
AssertTrue "File.Copy creates beta file", fso.FileExists(betaPhysical), betaPhysical

Set betaFile = fso.GetFile(betaPhysical)
betaFile.Name = "beta_renamed.txt"
AssertTrue "File.Name property rename updates file", fso.FileExists(betaRenamedPhysical), betaRenamedPhysical

betaFile.Move gammaPhysical
AssertTrue "File.Move creates gamma file", fso.FileExists(gammaPhysical), gammaPhysical
AssertTrue "File.Move removes previous beta name", (Not fso.FileExists(betaRenamedPhysical)), betaRenamedPhysical

Set gammaFile = fso.GetFile(gammaPhysical)
AssertEqual "Moved file keeps expected name", "gamma.txt", gammaFile.Name

WriteLine ""
WriteLine "SECTION | Folder object properties, collections, and methods"
Set nestedFolder = fso.CreateFolder(nestedPhysical)
Set baseFolder = fso.GetFolder(basePhysical)
AssertEqual "Folder.Name", "fso_runtime", baseFolder.Name
AssertTrue "Folder.ShortName is non-empty", (Len(CStr(baseFolder.ShortName)) > 0), CStr(baseFolder.ShortName)
AssertPathEqual "Folder.Path", basePhysical, baseFolder.Path
AssertTrue "Folder.ShortPath is non-empty", (Len(CStr(baseFolder.ShortPath)) > 0), CStr(baseFolder.ShortPath)
AssertTrue "Folder.Attributes is non-negative", (CLng(baseFolder.Attributes) >= 0), CStr(baseFolder.Attributes)
AssertTrue "Folder.DateCreated is non-empty", (Len(CStr(baseFolder.DateCreated)) > 0), CStr(baseFolder.DateCreated)
AssertTrue "Folder.DateLastAccessed is non-empty", (Len(CStr(baseFolder.DateLastAccessed)) > 0), CStr(baseFolder.DateLastAccessed)
AssertTrue "Folder.DateLastModified is non-empty", (Len(CStr(baseFolder.DateLastModified)) > 0), CStr(baseFolder.DateLastModified)
Set baseDrive = baseFolder.Drive
AssertEqual "Folder.Drive.DriveLetter", UCase(driveFromMethod.DriveLetter), UCase(baseDrive.DriveLetter)
AssertTrue "Folder.Size is non-negative", (CLng(baseFolder.Size) >= 0), CStr(baseFolder.Size)
AssertEqual "Folder.Type", "Folder", baseFolder.Type
AssertTrue "Folder.IsRootFolder is false", (Not CBool(baseFolder.IsRootFolder)), CStr(baseFolder.IsRootFolder)
Set baseParentFolder = baseFolder.ParentFolder
AssertPathEqual "Folder.ParentFolder.Path", testsPhysical, baseParentFolder.Path

Set folderTextStream = baseFolder.CreateTextFile("folder_created.txt", True)
folderTextStream.WriteLine "folder created"
folderTextStream.Close
AssertTrue "Folder.CreateTextFile creates file inside folder", fso.FileExists(folderCreatedPhysical), folderCreatedPhysical

folderFileCount = CLng(baseFolder.Files.Count)
folderSubFolderCount = CLng(baseFolder.SubFolders.Count)
AssertTrue "Folder.Files.Count reflects created files", (folderFileCount >= 3), CStr(folderFileCount)
AssertTrue "Folder.SubFolders.Count reflects nested folder", (folderSubFolderCount >= 1), CStr(folderSubFolderCount)

Set alphaFromCollection = baseFolder.Files.Item("stream_ops.txt")
AssertEqual "Folder.Files.Item returns stream fixture", "stream_ops.txt", alphaFromCollection.Name
Set nestedFromCollection = baseFolder.SubFolders.Item("nested")
AssertEqual "Folder.SubFolders.Item returns nested folder", "nested", nestedFromCollection.Name

nestedFolder.Name = "nested_renamed"
AssertTrue "Folder.Name property rename updates folder", fso.FolderExists(nestedRenamedPhysical), nestedRenamedPhysical

baseFolder.Copy copyPhysical, True
AssertTrue "Folder.Copy creates folder copy", fso.FolderExists(copyPhysical), copyPhysical

Set copiedFolder = fso.GetFolder(copyPhysical)
AssertTrue "Copied folder retains files", (CLng(copiedFolder.Files.Count) >= 3), CStr(copiedFolder.Files.Count)
AssertTrue "Copied folder retains subfolders", (CLng(copiedFolder.SubFolders.Count) >= 1), CStr(copiedFolder.SubFolders.Count)

copiedFolder.Move movedPhysical
AssertTrue "Folder.Move creates moved folder", fso.FolderExists(movedPhysical), movedPhysical
AssertTrue "Folder.Move removes original copy path", (Not fso.FolderExists(copyPhysical)), copyPhysical

Set movedFolder = fso.GetFolder(movedPhysical)
AssertEqual "Moved folder keeps expected name", "fso_runtime_moved", movedFolder.Name

WriteLine ""
WriteLine "SECTION | Root copy, move, and delete helpers"
fso.CopyFile streamPhysical, betaPhysical, True
AssertTrue "CopyFile creates beta file", fso.FileExists(betaPhysical), betaPhysical
fso.MoveFile betaPhysical, betaRenamedPhysical
AssertTrue "MoveFile moves beta file", fso.FileExists(betaRenamedPhysical), betaRenamedPhysical
AssertTrue "MoveFile clears old beta path", (Not fso.FileExists(betaPhysical)), betaPhysical
fso.DeleteFile betaRenamedPhysical
AssertTrue "DeleteFile removes beta renamed file", (Not fso.FileExists(betaRenamedPhysical)), betaRenamedPhysical

fso.CopyFolder basePhysical, copyPhysical, True
AssertTrue "CopyFolder creates a new folder copy", fso.FolderExists(copyPhysical), copyPhysical
fso.MoveFolder copyPhysical, movedPhysical
AssertTrue "MoveFolder creates moved folder", fso.FolderExists(movedPhysical), movedPhysical
AssertTrue "MoveFolder removes old copy folder", (Not fso.FolderExists(copyPhysical)), copyPhysical
fso.DeleteFolder movedPhysical
AssertTrue "DeleteFolder removes moved folder", (Not fso.FolderExists(movedPhysical)), movedPhysical

WriteLine ""
WriteLine "SECTION | Cleanup"
If fso.FolderExists(movedPhysical) Then
    fso.DeleteFolder movedPhysical
End If
If fso.FolderExists(copyPhysical) Then
    fso.DeleteFolder copyPhysical
End If
If fso.FolderExists(basePhysical) Then
    fso.DeleteFolder basePhysical
End If
AssertTrue "Cleanup removes base sandbox", (Not fso.FolderExists(basePhysical)), basePhysical
AssertTrue "Cleanup removes copy sandbox", (Not fso.FolderExists(copyPhysical)), copyPhysical
AssertTrue "Cleanup removes moved sandbox", (Not fso.FolderExists(movedPhysical)), movedPhysical

WriteLine ""
WriteLine "SUMMARY"
WriteLine "Passed: " & passCount
WriteLine "Failed: " & failCount
%>
