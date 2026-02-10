<!-- #include file="begin.asp"-->


<!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><form action=ad_customer.asp name=mainform method=post><%dim xCat
set xCat = Server.CreateObject("ADOX.Catalog")
set xCat.ActiveConnection = db.getConn
Response.Write "On Error Resume Next<br /><br />"
dim table,column, index, prop, createPrimary
for each table in xCat.Tables 
if lcase(left(table.name,2))="tb" then
Response.Write "db.execute(""CREATE TABLE [" & table.name & "]"")<br />"
'Response.Write "<u>Columns</u>: "& "<br/>"
createPrimary=false
for each column in table.columns
Response.Write "db.execute(""ALTER TABLE [" & table.name & "] ADD COLUMN [" & column.name & "] " & GetSQLTypeName(column) & """)<br />"
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
Response.Write "db.execute(""CREATE INDEX [PrimaryKey] ON " & table.name & "([iId]) WITH PRIMARY"")<br />"
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
Case 3	if field.Properties("Autoincrement").Value then GetSQLTypeName = "COUNTER NOT NULL" else GetSQLTypeName = "LONG NULL"
Case 7	GetSQLTypeName = "DATETIME NULL"
Case 11	GetSQLTypeName = "BIT NULL"
Case 6	GetSQLTypeName = "MONEY"
Case 128	GetSQLTypeName = "BINARY"
Case 17	GetSQLTypeName = "TINYINT"
Case 131	GetSQLTypeName = "DECIMAL"
Case 5	GetSQLTypeName = "FLOAT"
Case 2	GetSQLTypeName = "INTEGER NULL"
Case 4	GetSQLTypeName = "REAL"
Case 72	GetSQLTypeName = "UNIQUEIDENTIFIER"
Case 202	GetSQLTypeName = "TEXT(255)"
Case 203	GetSQLTypeName = "MEMO NULL"
End Select
End Function%></form><!-- #include file="ad_back.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
