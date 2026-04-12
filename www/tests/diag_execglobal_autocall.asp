<%
Function MakeObj()
  Dim d
  Set d = Server.CreateObject("Scripting.Dictionary")
  d.Add "id", "42"
  Set MakeObj = d
End Function
ExecuteGlobal "Function DynTest() : DynTest = MakeObj.Item(""id"") : End Function"
Response.Write DynTest()
%>
