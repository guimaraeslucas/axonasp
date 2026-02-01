<%
' Test database.rs property
On Error Resume Next

' Include asplite
<!-- #include file="../asplite/asplite.asp"-->

dim db : set db=aspL.plugin("database") : db.path="../db/sample.mdb"

dim rs : set rs=db.rs

Response.Write "rs object type: " & TypeName(rs) & "<br>"
Response.Write "rs is nothing: " & (rs is nothing) & "<br>"

if not (rs is nothing) then
    Response.Write "rs.EOF: " & rs.EOF & "<br>"
    Response.Write "rs.BOF: " & rs.BOF & "<br>"
    Response.Write "Attempting to open...<br>"
    rs.Open "select top 5 * from testdata"
    if Err.Number <> 0 then
        Response.Write "Error opening: " & Err.Description & "<br>"
    else
        Response.Write "Opened successfully<br>"
        Response.Write "rs.EOF after open: " & rs.EOF & "<br>"
        Response.Write "rs.RecordCount: " & rs.RecordCount & "<br>"
    end if
else
    Response.Write "rs is nothing - object was not created!<br>"
end if

if Err.Number <> 0 then
    Response.Write "Final Error: " & Err.Description & "<br>"
end if
%>
