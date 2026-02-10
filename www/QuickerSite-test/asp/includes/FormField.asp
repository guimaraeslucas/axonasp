
<%
private labelRadioRunner
labelRadioRunner=0
class cls_FormField
Public iId,iFormID,sName,sType,sValues,iRang,bMandatory,iMaxFileSize,iRows,iCols,iSize,iMaxlength,sRadioPlacement,bAutoResponder,sFileLocation,sAllowedExtensions,bUseForSending,bAllowMS
Public sPlaceholder
Private newRang,oldRang
Private Sub Class_Initialize
On Error Resume Next
iId=null
iRows=2
iCols=20
iSize=20
iMaxlength=255
iMaxFileSize=256
sType=sb_ff_text
sRadioPlacement=pl_Vertical
bMandatory=true
bAutoResponder=false
sAllowedExtensions=""
newRang=0
iFormID=convertLng(decrypt(request("iFormID")))
bUseForSending=false
bAllowMS=false
pick(decrypt(request("iFormFieldID")))
On Error Resume Next
end sub
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblFormField where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
iFormID	= rs("iFormID")
sName	= rs("sName")
sType	= rs("sType")
sValues	= rs("sValues")
iRang	= rs("iRang")
bMandatory	= rs("bMandatory")
iMaxFileSize	= rs("iMaxFileSize")
iMaxlength	= rs("iMaxlength")
iCols	= rs("iCols")
iRows	= rs("iRows")
iSize	= rs("iSize")
sRadioPlacement	= rs("sRadioPlacement")
bAutoResponder	= rs("bAutoResponder")
sFileLocation	= rs("sFileLocation")
sAllowedExtensions	= rs("sAllowedExtensions")
bUseForSending	= rs("bUseForSending")
bAllowMS	= rs("bAllowMS")
sPlaceholder=rs("sPlaceholder")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sType) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
end if
if (sType=sb_ff_select or sType=sb_ff_radio or sType=sb_ff_comment) and isLeeg(sValues) then
check=false
message.AddError("err_mandatory")
end if
if not isLeeg(sFileLocation) then
dim fso
set fso=server.CreateObject ("scripting.filesystemobject")
if convertBool(customer.bAllowStorageOutsideWWW) then
if not fso.folderExists(sFileLocation) then
check=false
message.AddError("err_foldernotexists")
end if
else
if instr(sFileLocation,server.mappath("/"))=0 then
check=false
message.AddError("err_foldernotallowed")
elseif not fso.folderExists(sFileLocation) then
check=false
message.AddError("err_foldernotexists")
end if
end if 
sFileLocation=replace(sFileLocation,"/","\",1,-1,1)
while right(sFileLocation,1)="\"
sFileLocation=left(sFileLocation,len(sFileLocation)-1)
wend
while instr(sFileLocation,"\\")<>0 
sFileLocation=replace(sFileLocation,"\\","\",1,-1,1)
wend 
if 	sFileLocation=server.MapPath (application("QS_CMS_userfiles")) then sFileLocation=""
end if
End Function
Public Function Save
dim rs
if check() then
save=true
else
save=false
exit function
end if
'set db=nothing
'set db=new cls_database
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblFormField where 1=2"
rs.AddNew
dim rangRS
set rangRS=db.execute("select count(*) from tblFormField where iFormID=" & iFormID)
iRang=convertGetal(clng(rangRS(0)))+1
else
rs.Open "select * from tblFormField where iId="& iId
end if
rs("iFormID")	= iFormID
rs("sName")	= sName
rs("sType")	= sType
rs("sValues")	= sValues
rs("iRang")	= iRang
rs("bMandatory")	= bMandatory
rs("iMaxFileSize")	= iMaxFileSize
rs("iMaxlength")	= iMaxlength
rs("iCols")	= iCols
rs("iRows")	= iRows
rs("iSize")	= iSize
rs("sRadioPlacement")	= sRadioPlacement
rs("bAutoResponder")	= bAutoResponder
rs("sFileLocation")	= sFileLocation
rs("sAllowedExtensions")	= sAllowedExtensions
rs("bUseForSending")	= bUseForSending
rs("bAllowMS")	= bAllowMS
rs("sPlaceholder")=sPlaceholder
dim setRangByUser
setRangByUser=false 
if convertGetal(newRang)<>0 and convertGetal(newRang)<>convertGetal(iRang) then
'set new rang set by user!
oldRang	= iRang
iRang	= newRang
rs("iRang")	= iRang
setRangByUser	= true
end if
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
if setRangByUser then
setRangByNumber oldRang,newRang
end if
if bUseForSending then
db.execute("update tblFormField set bUseForSending=" & getSQLBoolean(false) & " where iFormID="& iFormID & " and iId<>"& iId)
end if
end function
public function getRequestValues()
iFormID	= convertLng(decrypt(Request.Form ("iFormID")))
sType	= convertStr(Request.Form ("sType"))
sName	= convertStr(Request.Form ("sName"))
sValues	= Request.Form ("sValues")
bMandatory	= convertBool(Request.Form ("bMandatory"))
iMaxFileSize	= convertGetal(Request.form("iMaxFileSize"))
iMaxlength	= convertGetal(Request.form("iMaxlength"))
iCols	= convertGetal(Request.form("iCols"))
iRows	= convertGetal(Request.form("iRows"))
iSize	= convertGetal(Request.form("iSize"))
sRadioPlacement	= convertStr(Request.form("sRadioPlacement"))
bAutoResponder	= convertBool(Request.Form ("bAutoResponder"))
sFileLocation	= left(convertStr(Request.Form("sFileLocation")),255)
sAllowedExtensions	= lcase(left(convertStr(Request.Form("sAllowedExtensions")),255))
newRang	= convertGetal(Request.Form ("iRang"))
bUseForSending	= convertBool(Request.Form ("bUseForSending"))
bAllowMS	= convertBool(request.form("bAllowMS"))
sPlaceholder=convertStr(request.form("sPlaceholder"))

if iMaxLength=0 and sType<>sb_ff_textarea then iMaxLength=255
if iMaxFileSize=0 then iMaxFileSize=256
if iCols=0 then iCols=20
if iRows=0 then iRows=2
if iSize=0 then iSize=20
if sType=sb_ff_url or sType=sb_ff_email then
iMaxLength=1024
iCols=0
iRows=0
end if
if sType=sb_ff_date or sType=sb_ff_checkbox or sType=sb_ff_select or sType=sb_ff_file or sType=sb_ff_image or sType=sb_ff_radio or sType=sb_ff_hidden then
iSize=0
iMaxLength=0
iCols=0
iRows=0
end if 
if sType<>sType=sb_ff_file and sType<>sb_ff_image then
iMaxFileSize=0
end if
end function
public property get isMandatory
if bMandatory then 
isMandatory=l("yes")
else
isMandatory=l("no")
end if
end property
public function form
set form=new cls_form
form.pick(iFormID)
end function
public function moveUp
if isNumeriek(iId) then
if convertGetal(iRang)=1 then
exit function
else
db.execute("update tblFormField set iRang=iRang-1 where "& sqlNull("iFormID",iFormID) &" and iId="& iId)
db.execute("update tblFormField set iRang=iRang+1 where "& sqlNull("iFormID",iFormID) &" and iId<>"& iId & " and iRang=" & iRang-1)
end if
end if
end function
public function moveDown
if isNumeriek(iId) then
if convertGetal(iRang)=convertGetal(form.Fields.count) then
exit function
else
db.execute("update tblFormField set iRang=iRang+1 where "& sqlNull("iFormID",iFormID) &" and iId="& iId)
db.execute("update tblFormField set iRang=iRang-1 where "& sqlNull("iFormID",iFormID) &" and iId<>"& iId & " and iRang=" & iRang+1)
end if
end if
end function
private function setRangByNumber(oldR,newR)
if not isLeeg(iId) and convertGetal(oldR)<>convertGetal(newR) then
dim sql,rs
if oldR>newR then
sql="update tblFormField set iRang=iRang+1 where "
sql=sql&" iRang>=" & newR & " and iRang<" & oldR & " and iFormID="& iFormID
else
sql="update tblFormField set iRang=iRang-1 where "
sql=sql&" iRang>" & oldR & " and iRang<=" & newR & " and iFormID="& iFormID
end if
sql=sql&" and iId<>"& iId
set rs=db.execute(sql)
set rs=nothing
end if
end function
public function remove
if isNumeriek(iId) then
'bestanden verwijderen!
dim fso, rs, value
set fso=server.CreateObject ("scripting.filesystemobject")
if sType=sb_ff_file or sType=sb_ff_image then
set rs=db.execute("select sValue from tblFormFieldValue where iFormFieldID="&iId)
while not rs.eof
value=rs(0)
if isLeeg(sFileLocation) then	 
if fso.FileExists(server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & value)) then
fso.DeleteFile  server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & value)
end if
else
if fso.FileExists(sFileLocation & "\" & value) then
fso.DeleteFile  sFileLocation & "\" & value
end if
end if
rs.movenext
wend
set rs=nothing
end if
set fso=nothing
'rang updaten
set rs=db.execute("update tblFormField set iRang=iRang-1 where "& sqlNull("iFormID",iFormID) &" and iRang>"& iRang)
set rs=nothing
'waarden verwijderen
set rs=db.execute("delete from tblFormFieldValue where iFormFieldID="& iId)
set rs=nothing
'record verwijderen
set rs=db.execute("delete from tblFormField where iId="& iId)
set rs=nothing
end if
end function
Public Function showSelected(selected)
dim listArr, i, list
set list=server.CreateObject ("scripting.dictionary")
listArr=split(sValues,vbcrlf)
for i=lbound(listArr) to ubound(listArr)
On Error Resume Next
list.Add listArr(i),""
On Error Goto 0
next
Dim key
For each key in list
showSelected = showSelected & "<option value=" & """" & quotRep(key) & """"
If selected = key Then
showSelected = showSelected & " selected='selected'"
End If
showSelected = showSelected & ">" & quotRep(key) & "</option>"
Next
End Function



Public Function showSelectedRadio(selected)
dim listArr, i, list
set list=server.CreateObject ("scripting.dictionary")
listArr=split(sValues,vbcrlf)
for i=lbound(listArr) to ubound(listArr)
On Error Resume Next
list.Add listArr(i),""
On Error Goto 0
next
Dim key
For each key in list
labelRadioRunner=labelRadioRunner+1
showSelectedRadio = showSelectedRadio & "<input "

if bMandatory then
	showSelectedRadio = showSelectedRadio & " required "
end if

showSelectedRadio = showSelectedRadio & " type='radio'"
If convertStr(quotRep(selected))=convertStr(quotRep(key)) Then
showSelectedRadio = showSelectedRadio & " checked='checked' "
End If
showSelectedRadio = showSelectedRadio & " name='" & encrypt(iId) & "' id=""qsRadioID" & labelRadioRunner & """ value="& """" & quotRep(key) & """"
showSelectedRadio = showSelectedRadio & " /> <label for=""qsRadioID" & labelRadioRunner & """>" & key & "</label>"
if sRadioPlacement=pl_Vertical then
showSelectedRadio = showSelectedRadio & "<br />"
else
showSelectedRadio = showSelectedRadio & "&nbsp;"
end if
Next
End Function
Public Function showSelectedRadioCB(selected)
selected=selected&", "
dim selectedValues
set selectedValues=server.createObject("scripting.dictionary")
selected=split(selected,"_QSDELIMITER, ")
for i=lbound(selected) to ubound(selected)
if not selectedValues.exists(selected(i)) then
selectedValues.Add trim(convertStr(selected(i))),""
end if
next
dim listArr, i, list
set list=server.CreateObject ("scripting.dictionary")
listArr=split(sValues,vbcrlf)
for i=lbound(listArr) to ubound(listArr)
On Error Resume Next
list.Add listArr(i),""
On Error Goto 0
next
Dim key
For each key in list
labelRadioRunner=labelRadioRunner+1
'response.write selectedValues.count
showSelectedRadioCB = showSelectedRadioCB & "<input "

if bMandatory then
	showSelectedRadioCB = showSelectedRadioCB & " required "
end if

showSelectedRadioCB = showSelectedRadioCB & " type=""checkbox"""
If selectedValues.exists(convertStr(key)) Then
showSelectedRadioCB = showSelectedRadioCB & " checked=""checked"" "
End If
showSelectedRadioCB = showSelectedRadioCB & " name=""" & encrypt(iId) & """ id=""qsRadioID" & labelRadioRunner & """ value="& """" & quotRep(key) & "_QSDELIMITER"""
showSelectedRadioCB = showSelectedRadioCB & " /> <label for=""qsRadioID" & labelRadioRunner & """>"& key & "</label>" 
if sRadioPlacement=pl_Vertical then
showSelectedRadioCB = showSelectedRadioCB & "<br />"
else
showSelectedRadioCB = showSelectedRadioCB & "&nbsp;"
end if
Next
End Function
public function copy(formID)
if isNumeriek(iId) then
iFormID=formID
iId=null
save()
end if
end function
public property get sFileURL(mode)
end property
public property get form
set form=new cls_form
if convertGetal(iFormID)<>0 then
form.pick(iFormID)
end if
end property
end class%>
