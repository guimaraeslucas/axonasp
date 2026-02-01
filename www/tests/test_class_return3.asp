<%
Class TestClass
    Public Function GetRS()
        Set GetRS = Server.CreateObject("adodb.recordset")
        GetRS.CursorType = 1
        GetRS.LockType = 3
        Set GetRS.ActiveConnection = Server.CreateObject("adodb.connection")
    End Function
End Class

dim tc : set tc = new TestClass
dim rs : set rs = tc.GetRS()
Response.Write "TypeName: " & typename(rs) & "<br>"
%>
