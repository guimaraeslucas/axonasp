<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AxonASP - System and Runtime Custom Functions</title>
</head>
<body>
<h1>System and Runtime Custom Functions</h1>
<pre>
<%
Dim originalWd, tempDir, testFile, testLink
Dim envItems, i, maxPreview

Response.Write "=== OS-like Ax Functions ===" & vbCrLf
Response.Write "AxHostNameValue(): " & AxHostNameValue() & vbCrLf
Response.Write "AxProcessId(): " & AxProcessId() & vbCrLf
Response.Write "AxEffectiveUserId(): " & AxEffectiveUserId() & vbCrLf
Response.Write "AxCurrentDir(): " & AxCurrentDir() & vbCrLf
Response.Write "AxGetEnv('PATH') length: " & Len(AxGetEnv("PATH")) & vbCrLf
Response.Write "AxEnvironmentValue('DOES_NOT_EXIST', 'fallback'): " & AxEnvironmentValue("DOES_NOT_EXIST", "fallback") & vbCrLf
Response.Write "AxUserHomeDirPath(): " & AxUserHomeDirPath() & vbCrLf
Response.Write "AxUserConfigDirPath(): " & AxUserConfigDirPath() & vbCrLf
Response.Write "AxUserCacheDirPath(): " & AxUserCacheDirPath() & vbCrLf
Response.Write "AxDirSeparator(): " & AxDirSeparator() & vbCrLf
Response.Write "AxPathListSeparator(): " & AxPathListSeparator() & vbCrLf
Response.Write "AxIsPathSeparator('/'): " & AxIsPathSeparator("/") & vbCrLf
Response.Write "AxIsPathSeparator('a'): " & AxIsPathSeparator("a") & vbCrLf
Response.Write "AxIntegerSizeBytes(): " & AxIntegerSizeBytes() & vbCrLf
Response.Write "AxIntegerMax(): " & AxIntegerMax() & vbCrLf
Response.Write "AxIntegerMin(): " & AxIntegerMin() & vbCrLf
Response.Write "AxFloatPrecisionDigits(): " & AxFloatPrecisionDigits() & vbCrLf
Response.Write "AxSmallestFloatValue(): " & AxSmallestFloatValue() & vbCrLf
Response.Write "AxPlatformBits(): " & AxPlatformBits() & vbCrLf
Response.Write "AxExecutablePath(): " & AxExecutablePath() & vbCrLf

envItems = AxEnvironmentList()
maxPreview = 5
Response.Write "AxEnvironmentList() preview:" & vbCrLf
For i = 0 To UBound(envItems)
    Response.Write "  - " & envItems(i) & vbCrLf
    maxPreview = maxPreview - 1
    If maxPreview <= 0 Then Exit For
Next

originalWd = AxCurrentDir()
tempDir = Server.MapPath("/tests/temp")
testFile = Server.MapPath("/tests/test.txt")
testLink = tempDir & AxDirSeparator() & "ax_custom_func_tmp.link"

Response.Write "AxChangeDir(tempDir): " & AxChangeDir(tempDir) & vbCrLf
Response.Write "AxCurrentDir() after change: " & AxCurrentDir() & vbCrLf
Response.Write "AxChangeDir(originalWd): " & AxChangeDir(originalWd) & vbCrLf
Response.Write "AxChangeTimes(testFile, 1700000000, 1700000001): " & AxChangeTimes(testFile, 1700000000, 1700000001) & vbCrLf
Response.Write "AxChangeMode(testFile, '0644'): " & AxChangeMode(testFile, "0644") & vbCrLf
Response.Write "AxCreateLink(testFile, testLink): " & AxCreateLink(testFile, testLink) & " (may be false on restricted environments)" & vbCrLf
Response.Write "AxChangeOwner(testFile, 0, 0): " & AxChangeOwner(testFile, 0, 0) & " (expected false on Windows/non-privileged environments)" & vbCrLf

Response.Write vbCrLf & "=== Runtime Functions ===" & vbCrLf
Response.Write "AxLastModified(): " & AxLastModified() & vbCrLf
Response.Write "AxSystemInfo('a'): " & AxSystemInfo("a") & vbCrLf
Response.Write "AxSystemInfo('s'): " & AxSystemInfo("s") & vbCrLf
Response.Write "AxSystemInfo('n'): " & AxSystemInfo("n") & vbCrLf
Response.Write "AxSystemInfo('r'): " & AxSystemInfo("r") & vbCrLf
Response.Write "AxSystemInfo('v'): " & AxSystemInfo("v") & vbCrLf
Response.Write "AxSystemInfo('m'): " & AxSystemInfo("m") & vbCrLf
Response.Write "AxCurrentUser(): " & AxCurrentUser() & vbCrLf
Response.Write "AxVersion(): " & AxVersion() & vbCrLf
Response.Write "AxVersion('version_id'): " & AxVersion("version_id") & vbCrLf
Response.Write vbCrLf & "AxRuntimeInfo() output below:" & vbCrLf
Call AxRuntimeInfo()

Response.Write vbCrLf & "=== Done ===" & vbCrLf
%>
</pre>
</body>
</html>
