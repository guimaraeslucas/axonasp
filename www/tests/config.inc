<%
' config.inc - Sample configuration file for testing AxIncludeOnce
Dim g_AppName
Dim g_AppVersion
Dim g_IsConfigLoaded

g_AppName = "Test Application"
g_AppVersion = "1.0"
g_IsConfigLoaded = True

Response.Write " Config loaded at " & Now() & ">" & vbCrLf
%>

