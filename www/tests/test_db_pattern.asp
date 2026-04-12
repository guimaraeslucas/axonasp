<%
' Test function returning recordset-like object
On Error Resume Next

' Mock database class
Class MockConnection
    Public Function Execute(sql)
        ' Return a mock recordset
        Dim mock
        Set mock = New MockRecordset
        mock.SQL = sql
        Set Execute = mock
    End Function
End Class

Class MockRecordset
    Public SQL
    Public Name
    Function Item(field)
        Item = "result_" & field
    End Function
End Class

Function GetConnection()
    Dim conn
    Set conn = New MockConnection
    Set GetConnection = conn
End Function

Dim db
Set db = GetConnection()

' This is the pattern from QuickerSite
Dim rs, sql
sql = "select * from table where id=" & 1
Response.Write "sql = " & sql & vbCrLf

Set rs = db.Execute(sql)
Response.Write "rs.SQL = " & rs.SQL & vbCrLf
Response.Write "rs('name') = " & rs("name") & vbCrLf
%>
