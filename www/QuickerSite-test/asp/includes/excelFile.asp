
<%Class cls_excelFile
private pw
private quote
private fso
sub class_initialize
quote="IEExport_"
pw=GeneratePassWord()
set fso=server.CreateObject ("scripting.filesystemobject")
end sub
sub class_terminate
set fso=nothing
end sub
Public Function export(table)
on error resume next
table=convertTo_iso_8859_1(prepareForExport(table))
deleteOthers()
dim file
set file=fso.CreateTextFile(server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & quote & pw & ".xls"),1,-1)
file.writeline "<html xmlns:o=" & """" & "urn:schemas-microsoft-com:office:office" & """" &" xmlns:x=" & """" & "urn:schemas-microsoft-com:office:excel" & """" & " xmlns=" & """" & "http://www.w3.org/TR/REC-html40" & """" & ">"
file.writeline "<head>"
file.writeline "<meta HTTP-EQUIV='Content-Type' content='text/html; charset="& QS_CHARSET &"' />"
file.writeline "<title>Export</title>"
file.writeline "<style>br {mso-data-placement:same-cell;}tr {vertical-align:top;}</style></head><body>"
file.writeline table
file.writeline "</body></html>"
set file=nothing
on error goto 0
End Function
public property get downloadLink
downloadLink="<a class=""art-button"" href="& """" & C_VIRT_DIR & Application("QS_CMS_userfiles") & quote & pw & ".xls" & """" & " target='p"&pw&"'>"& l("downloadexcel") &"</a>"
end property
public property get filename
filename=quote & pw
end property
public sub  redirectLink
Response.Redirect (C_VIRT_DIR & Application("QS_CMS_userfiles") & filename & ".xls")
end sub
private sub deleteOthers
on error resume next
dim folder
set folder=fso.GetFolder (server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles")))
dim files
set files=folder.files
dim fileKey
for each fileKey in files
if instr(fileKey.name,quote)<>0 then
fileKey.delete()
end if 
next 
set folder=nothing
on error goto 0
end sub
End class%>
