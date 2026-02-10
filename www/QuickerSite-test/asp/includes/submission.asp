
<%class cls_submission
Public iId
Public iFormID
Public iItemID
Public dCreatedTS
Public dUpdatedTS
Public sVisitorDetails
Private Sub Class_Initialize
On Error Resume Next
iId=null
pick(decrypt(request("iSubmissionID")))
On Error Resume Next
end sub
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblFormSubmission where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
iFormID	= rs("iFormID")
iItemID	= rs("iItemID")
dCreatedTS	= rs("dCreatedTS")
dUpdatedTS	= rs("dUpdatedTS")
sVisitorDetails	= rs("sVisitorDetails")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
End Function
Public Function Save
sVisitorDetails=getVisitorDetails()
dim rs
if check() then
save=true
else
save=false
exit function
end if
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblFormSubmission where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblFormSubmission where iId="& iId
end if
rs("iFormID")	= iFormID
rs("iItemID")	= iItemID
rs("dUpdatedTS")	= now()
rs("sVisitorDetails")	= sVisitorDetails
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function form
set form=new cls_form
form.pick(iFormID)
end function
public function item
set item=new cls_catalogItem
item.pick(iItemID)
end function
public function remove(fields)
if isNumeriek(iId) then
dim fso
set fso=server.CreateObject ("scripting.filesystemobject")
dim cValues,vKey
set cValues=values(fields)
for each vKey in cValues
if convertGetal(fields(vKey).sType)=convertGetal(sb_ff_file) or convertGetal(fields(vKey).sType)=convertGetal(sb_ff_image) then
if isLeeg(fields(vKey).sFileLocation) then
if fso.FileExists(server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") &cValues(vKey))) then
fso.DeleteFile  server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") &cValues(vKey))
end if
else
if fso.FileExists(fields(vKey).sFileLocation & "\" &cValues(vKey)) then
fso.DeleteFile  fields(vKey).sFileLocation & "\" &cValues(vKey)
end if
end if
end if
next
dim rs
set rs=db.execute("delete from tblFormFieldValue where iSubmissionID="& iId)
set rs=nothing
set rs=db.execute("delete from tblFormSubmission where iId="& iId)
set rs=nothing
end if
end function
public function values(fields)
set values=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, iFormFieldId, sValue
set rs=db.execute("select iFormFieldId,sValue from tblFormFieldValue where iSubmissionID="& iId)
while not rs.eof
iFormFieldId	= rs(0)
sValue	= rs(1)
values.Add convertGetal(iFormFieldId),sValue
rs.movenext
wend
set rs=nothing
end if
end function


function sCopyValues

	sCopyValues=""

	if isNumeriek(iId) then
	
		dim cfields,cf,cvalues
		set cfields=form.fields
		set cvalues=values(cfields)
		
		for each cf in cfields
			sCopyValues=sCopyValues & "<tr><td style=""font-style:italic;text-align:right"">" & sanitize(cfields(cf).sName) & "</td><td>" & sanitize(cvalues(cf)) & "</td></tr>"
		next
		
		set cfields=nothing		
		
		
	end if
	
	if sCopyValues<>"" then
	
		sCopyValues="<table cellpadding=""5"" cellspacing=""0"" border=""1"">" & sCopyValues & "</table>"
	
	end if
	
	'response.write sCopyValues
	'response.end 

end function

end class%>
