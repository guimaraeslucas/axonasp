<!-- #include file="begin.asp"-->


<!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><form action=ad_customer.asp name=mainform method=post><%dim xCat
set xCat = Server.CreateObject("ADOX.Catalog")
set xCat.ActiveConnection = db.getConn
Response.Write "On Error Resume Next<br /><br />"
dim table,column, index, prop, createPrimary
for each table in xCat.Tables 
if lcase(left(table.name,2))="tb" then
Response.Write "db.execute(""CREATE TABLE [" & table.name & "] (QS_ID int NULL)"")<br />"
'Response.Write "<u>Columns</u>: "& "<br/>"
createPrimary=false
for each column in table.columns
Response.Write "db.execute(""ALTER TABLE [" & table.name & "] ADD [" & column.name & "] " & GetSQLTypeName(column) & """)<br />"
'Response.Write "if err.number<>0 then <br /> err.clear() <br /> response.write """& table.name  & " - " & column.name & "&lt;br&gt;&lt;br&gt;""<br />end if<br /><br />" & "<br /><br />"
if lcase(column.name)="iid" then
createPrimary=true
end if
for each prop in column.properties
'Response.Write "Prop Name:" & prop.name & " - "
'Response.Write "Prop Type:" & prop.Type & " - "
'Response.Write "Prop Value:" & prop.Value & "<br/>"
next
next
if createPrimary then
'Response.Write "db.execute(""CREATE INDEX [PrimaryKey] ON " & table.name & "([iId]) WITH PRIMARY"")<br />"
end if
Response.write "<br /><br />"
'Response.Write "<u>Indexes</u>: "& "<br/>"
for each index in table.indexes
'	Response.Write "Name: " & index.name& "<br/>"
'	Response.Write "Primary key: " & index.PrimaryKey & "<br/>"
next
end if
next
Response.Write "On Error Goto 0<br /><br />"
Function GetSQLTypeName(Field)
Select Case field.type
Case 3	if field.Properties("Autoincrement").Value then GetSQLTypeName = " INT IDENTITY(1,1) NOT NULL" else GetSQLTypeName = "INT NULL"
Case 7	GetSQLTypeName = "datetime NULL"
Case 11	GetSQLTypeName = "bit NULL"
Case 6	GetSQLTypeName = "MONEY"
Case 128	GetSQLTypeName = "BINARY"
Case 17	GetSQLTypeName = "int"
Case 131	GetSQLTypeName = "int"
Case 5	GetSQLTypeName = "int"
Case 2	GetSQLTypeName = "int NULL"
Case 4	GetSQLTypeName = "REAL"
Case 72	GetSQLTypeName = "UNIQUEIDENTIFIER"
Case 202	GetSQLTypeName = "nvarchar(255)"
Case 203	GetSQLTypeName = "ntext NULL"
End Select
End Function%></form><!-- #include file="ad_back.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
