
<%dim db
set db = new cls_database
Class cls_Database
private p_getConn
private p_getConnLabels
Private Sub Class_Initialize()
set p_getConn = nothing
set p_getConnLabels = nothing
End Sub
Public Function Execute(sql)
On Error Resume Next
dim connection
set connection = getConn()
Set Execute = connection.Execute(sql)
On Error Goto 0
End Function
Public Function ExecuteLabels(sql)
On Error Resume Next
dim connection
set connection = getConnLabels()
Set ExecuteLabels = connection.Execute(sql)
On Error Goto 0
End Function
Public Function GetDynamicRS
On Error Resume Next
Set GetDynamicRS = server.CreateObject ("adodb.recordset")
GetDynamicRS.CursorType = 1
GetDynamicRS.LockType = 3
set GetDynamicRS.ActiveConnection = getConn()
On Error Goto 0
End Function
Public Function GetDynamicRSLabels
On Error Resume Next
Set GetDynamicRSLabels = server.CreateObject ("adodb.recordset")
GetDynamicRSLabels.CursorType = 1
GetDynamicRSLabels.LockType = 3
GetDynamicRSLabels.ActiveConnection = getConnLabels()
On Error Goto 0
End Function
Public function getConn()
On Error Resume Next
if p_getConn is nothing then
Set p_getConn = Server.Createobject("ADODB.Connection")
select case QS_DBS
case 1 'Access             
if not dataIsSafe then
dumpwarning
end if
p_getConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&C_DATABASE
case 2 'SQL Server
if isLeeg(SQL2005_UID) then
p_getConn.Open "Provider=SQLOLEDB;SERVER=" & SQL2005_SERVER & ";initial catalog=" & SQL2005_DB  & ";Integrated Security=SSPI;"
else
p_getConn.Open "Provider=SQLOLEDB;SERVER=" & SQL2005_SERVER & ";initial catalog=" & SQL2005_DB  & ";User ID=" & SQL2005_UID & ";Password=" & SQL2005_PWD
end if
case 3 'MySQL
p_getConn.Open "Driver={MySQL ODBC 3.51 Driver};Server="&MySQL_SERVER&";Port="&MySQL_PORT&";Database="&MySQL_DB&";User="&MySQL_UID&";Password="&MySQL_PWD&";Option="&MySQL_OPTION&";"
end select
end if
set getConn=p_getConn
On Error Goto 0
end function
Public function getConnLabels()
On Error Resume Next
if p_getConnLabels is nothing then
Set p_getConnLabels = Server.Createobject("ADODB.Connection")
p_getConnLabels.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&C_DATABASE_LABELS
end if
set getConnLabels=p_getConnLabels
On Error Goto 0
end function
Private function dataIsSafe
if right(trim(lcase(C_DATABASE)),19)="\db\quickersite.mdb" and not devVersion then
dataIsSafe=false
else
dataIsSafe=true
end if
end function
private function dumpwarning
    
        response.clear()
        err.clear()
        on error resume next
        dim fso,dbFile,dbNewName,wcFile,wcContent,wcNewContent,dbLNewName
        set fso=server.createobject ("scripting.filesystemobject")
        set dbFile=fso.getFile (C_DATABASE)
        dbNewName="data_" & lcase(generatePassword()) & ".mdb"
        dbFile.name=dbNewName        
        set dbFile=nothing
        set dbFile=fso.getFile (C_DATABASE_LABELS)
        dbLNewName="labels_" & lcase(generatePassword()) & ".mdb"
        dbFile.name=dbLNewName       
        set dbFile=nothing
        'change web_config
        set wcFile=fso.OpenTextFile(server.mappath(C_DIRECTORY_QUICKERSITE & "/asp/config/web_config.asp"), 1, False) 
        wcContent=wcFile.readAll       
        set wcFile=nothing
        wcNewContent=replace(wcContent,"QuickerSite.mdb",dbNewName,1,-1,1)
        wcNewContent=replace(wcNewContent,"QuickerLabels.mdb",dbLNewName,1,-1,1)
                
        Set wcContent = fso.CreateTextFile(server.mappath(C_DIRECTORY_QUICKERSITE & "/asp/config/web_config.asp"), True)
        wcContent.WriteLine(wcNewContent)       
        set wcContent=nothing
        if err.number=0 then
            set fso=nothing 
response.redirect(C_DIRECTORY_QUICKERSITE & "/default.asp") 
        else
dim dbNewNameReset,dbLNewNameReset
'rewind everything!
set dbFile=fso.getFile (C_DATABASE)
dbNewNameReset="QuickerSite.mdb"
dbFile.name=dbNewNameReset        
set dbFile=nothing
set dbFile=fso.getFile (C_DATABASE_LABELS)
dbLNewNameReset="QuickerLabels.mdb"
dbFile.name=dbLNewNameReset       
set dbFile=nothing
set wcFile=fso.OpenTextFile(server.mappath(C_DIRECTORY_QUICKERSITE & "/asp/config/web_config.asp"), 1, False) 
wcContent=wcFile.readAll       
set wcFile=nothing
wcNewContent=replace(wcContent,dbNewName,"QuickerSite.mdb",1,-1,1)
wcNewContent=replace(wcNewContent,dbLNewName,"QuickerLabels.mdb",1,-1,1)
Set wcContent = fso.CreateTextFile(server.mappath(C_DIRECTORY_QUICKERSITE & "/asp/config/web_config.asp"), True)
wcContent.WriteLine(wcNewContent)       
set wcContent=nothing
set fso=nothing 
            Response.Redirect C_DIRECTORY_QUICKERSITE & "/asp/warning.htm"
        end if
        on error goto 0     
end function
End Class%>
