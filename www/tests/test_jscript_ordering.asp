<%@ Language="VBScript" %>
<%
' Call BEFORE definition - like test_aspcheck.asp does
Response.Write "=== Calling getJVer() BEFORE script block ===" & vbCrLf
Response.Write "getJVer result: [" & getjver() & "]" & vbCrLf
Response.Write vbCrLf
%>
<script runat="server" language="JScript">
    function getJVer() {
        return ScriptEngineMajorVersion() + "." + ScriptEngineMinorVersion() + "." + ScriptEngineBuildVersion();
    }
</script>
<%
' Call AFTER definition
Response.Write "=== Calling getJVer() AFTER script block ===" & vbCrLf
Response.Write "getJVer result: [" & getjver() & "]" & vbCrLf
%>