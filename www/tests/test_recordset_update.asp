<%
Option Explicit

Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Driver={SQLite3};Data Source=:memory:"

conn.Execute "CREATE TABLE test (id INTEGER, name TEXT)"
conn.Execute "INSERT INTO test VALUES (1, 'OldValue')"

Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")

' Test 1: Update Existing (Must use WHERE for our implementation to know what to update)
rs.Open "SELECT * FROM test WHERE id=1", conn

If rs.EOF Then
    Response.Write "ERROR: No records found."
Else
    Response.Write "Current Value: " & rs("name") & "<br>"
    rs("name") = "NewValue"
    Response.Write "Value after assignment: " & rs("name") & "<br>"
    rs.Update
    Response.Write "Update called.<br>"
End If
rs.Close

' Verify DB
Dim rsCheck
Set rsCheck = Server.CreateObject("ADODB.Recordset")
rsCheck.Open "SELECT * FROM test WHERE id=1", conn

If Not rsCheck.EOF Then
    Response.Write "Value from DB after Update: " & rsCheck("name") & "<br>"
    If rsCheck("name") = "NewValue" Then
        Response.Write "PASS: Database update worked.<br>"
    Else
        Response.Write "FAIL: Database update failed.<br>"
    End If
End If
rsCheck.Close

' Test 2: AddNew
Dim rsAdd
Set rsAdd = Server.CreateObject("ADODB.Recordset")
rsAdd.Open "SELECT * FROM test WHERE 1=2", conn
rsAdd.AddNew
rsAdd("id") = 2
rsAdd("name") = "AddedValue"
rsAdd.Update
rsAdd.Close

' Verify DB Add
Dim rsCheckAdd
Set rsCheckAdd = Server.CreateObject("ADODB.Recordset")
rsCheckAdd.Open "SELECT * FROM test WHERE id=2", conn

If Not rsCheckAdd.EOF Then
    Response.Write "Value from DB after AddNew: " & rsCheckAdd("name") & "<br>"
    If rsCheckAdd("name") = "AddedValue" Then
        Response.Write "PASS: Database AddNew worked.<br>"
    Else
        Response.Write "FAIL: Database AddNew failed (Wrong value).<br>"
    End If
Else
    Response.Write "FAIL: Database AddNew failed (Record not found).<br>"
End If
rsCheckAdd.Close

conn.Close
%>
