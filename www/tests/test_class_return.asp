<%
Class TestClass
    Public Function GetRS()
        Set GetRS = Server.CreateObject("adodb.recordset")
    End Function
End Class

dim tc : set tc = new TestClass
dim rs : set rs = tc.GetRS()
Response.Write "TypeName: " & typename(rs) & "<br>"
%>
