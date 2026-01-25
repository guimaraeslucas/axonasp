<%
Response.Write "Testing ScriptEngine functions<br/>"

Dim engine
engine = ScriptEngine()
Response.Write "ScriptEngine() returned: [" & engine & "]<br/>"
Response.Write "Type: " & TypeName(engine) & "<br/>"

Dim buildVer
buildVer = ScriptEngineBuildVersion()
Response.Write "ScriptEngineBuildVersion() returned: [" & buildVer & "]<br/>"
Response.Write "Type: " & TypeName(buildVer) & "<br/>"

Dim majorVer
majorVer = ScriptEngineMajorVersion()
Response.Write "ScriptEngineMajorVersion() returned: [" & majorVer & "]<br/>"
Response.Write "Type: " & TypeName(majorVer) & "<br/>"

Dim minorVer
minorVer = ScriptEngineMinorVersion()
Response.Write "ScriptEngineMinorVersion() returned: [" & minorVer & "]<br/>"
Response.Write "Type: " & TypeName(minorVer) & "<br/>"

Response.Write "<hr/>"
Response.Write "Testing Eval function<br/>"

Dim evalRes
evalRes = Eval("42")
Response.Write "Eval('42') returned: [" & evalRes & "]<br/>"
Response.Write "Type: " & TypeName(evalRes) & "<br/>"
%>
