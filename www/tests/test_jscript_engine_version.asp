<%@ Language="VBScript" %>
<%
Response.Write "=== VBScript Engine Info ===" & vbCrLf
Response.Write "ScriptEngine: " & ScriptEngine & vbCrLf
Response.Write "Major: " & ScriptEngineMajorVersion & vbCrLf
Response.Write "Minor: " & ScriptEngineMinorVersion & vbCrLf
Response.Write "Build: " & ScriptEngineBuildVersion & vbCrLf
Response.Write vbCrLf
%>
<script runat="server" language="JScript">
    Response.Write("=== JScript Engine Info ===\n");
    Response.Write("ScriptEngine: " + ScriptEngine() + "\n");
    Response.Write("Major: " + ScriptEngineMajorVersion() + "\n");
    Response.Write("Minor: " + ScriptEngineMinorVersion() + "\n");
    Response.Write("Build: " + ScriptEngineBuildVersion() + "\n");
    Response.Write("\n");

    function getJVer() {
        return ScriptEngineMajorVersion() + "." + ScriptEngineMinorVersion() + "." + ScriptEngineBuildVersion();
    }
</script>
<%
Response.Write "=== Calling getJVer() from VBScript ===" & vbCrLf
Response.Write "getJVer result: " & getjver() & vbCrLf
%>