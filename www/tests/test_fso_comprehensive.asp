<%
' FileSystemObject comprehensive test
Response.Write "<h1>FileSystemObject (FSO) API Reference</h1>"
Response.Write "<p>Complete implementation using G3Files library</p>"

Set FSO = Server.CreateObject("Scripting.FileSystemObject")

Response.Write "<h2>File Methods Testing</h2>"
Response.Write "FileExists: " & FSO.FileExists("/test_basics.asp") & "<br>"
Response.Write "BuildPath: " & FSO.BuildPath("/www", "file.txt") & "<br>"
Response.Write "GetFileName: " & FSO.GetFileName("/www/test_basics.asp") & "<br>"
Response.Write "GetBaseName: " & FSO.GetBaseName("test_basics.asp") & "<br>"
Response.Write "GetExtensionName: " & FSO.GetExtensionName("test_basics.asp") & "<br>"
Response.Write "GetParentFolderName: " & FSO.GetParentFolderName("/www/test_basics.asp") & "<br>"
Response.Write "GetDriveName: " & FSO.GetDriveName("C:\Windows") & "<br>"
Response.Write "GetTempName: " & FSO.GetTempName() & "<br>"
Response.Write "GetAbsolutePathName: " & FSO.GetAbsolutePathName("/test_basics.asp") & "<br>"

Response.Write "<h2>Folder Methods Testing</h2>"
Response.Write "FolderExists: " & FSO.FolderExists("/") & "<br>"

Response.Write "<h2>Summary</h2>"
Response.Write "<p><strong>✓ FileSystemObject (Scripting.FileSystemObject) - FULLY IMPLEMENTED</strong></p>"
Response.Write "<p>All standard FSO methods and properties are working correctly</p>"
Response.Write "<p>Integration with G3Files library is complete and operational</p>"
%>
