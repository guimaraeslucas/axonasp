<%
' Test with object interaction
Dim dict
Set dict = Server.CreateObject(""Scripting.Dictionary"")
dict.Add ""key1"", ""value1""

Response.Write ""Dict(key1) via direct: "" & dict(""key1"") & vbCrLf

' Eval should access the dict variable
Response.Write ""Dict(key1) via Eval: "" & Eval(""dict(""""key1"""")""") & vbCrLf

Response.Write ""Object tests passed!"" & vbCrLf
%>
