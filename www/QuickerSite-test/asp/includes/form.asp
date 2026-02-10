
<%class cls_form
Public iId, sName, sComment, sCode, sIntro, sFeedback, sTo, sSubject, sRedirect, sRedirectPrefix, sButton, sReset, iCustomerID, dUpdatedTS
Public dCreatedTS, bSendEmail, bCookie, bAttachFiles, bCaptcha, sQAalign, bAutoResponder, sAutoResponse, sAutoResponseSubject, sAutoResponseFromName
Public sAutoResponseFromEmail, sScriptUponSubmission, sAutoResponseWebmaster, overruleCID
Private resetJS, fromAddress
Private Sub Class_Initialize
On Error Resume Next
iId	= null
bSendEmail	= false
bCookie	= false
bAttachFiles	= false
bCaptcha	= false
bAutoResponder	= false
sQAalign	= QS_QleftAright
sButton	= l("send")
sReset	= l("reset")
overruleCID	= null
pick(decrypt(request("iFormID")))
if isLeeg(iId) then 
sTo = customer.webmasterEmail
sAutoResponseFromEmail = customer.webmasterEmail
sAutoResponseFromName = customer.sitename
end if
On Error Goto 0
end sub
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblForm where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
sName	= rs("sName")
sComment	= rs("sComment")
sCode	= rs("sCode")
sIntro	= rs("sIntro")
sFeedback	= rs("sFeedback")
sTo	= rs("sTo")
sSubject	= rs("sSubject")
sRedirect	= rs("sRedirect")
sRedirectPrefix	= rs("sRedirectPrefix")
sButton	= rs("sButton")
sReset	= rs("sReset")
iCustomerID	= rs("iCustomerID")
dUpdatedTS	= rs("dUpdatedTS")
dCreatedTS	= rs("dCreatedTS")
bSendEmail	= rs("bSendEmail")
bCookie	= rs("bCookie")
bAttachFiles	= rs("bAttachFiles")
bCaptcha	= rs("bCaptcha")
sQAalign	= rs("sQAalign")
bAutoResponder	= rs("bAutoResponder")
sAutoResponse	= rs("sAutoResponse")
sAutoResponseSubject	= rs("sAutoResponseSubject")
sAutoResponseFromName	= rs("sAutoResponseFromName")
sAutoResponseFromEmail	= rs("sAutoResponseFromEmail")
sScriptUponSubmission	= rs("sScriptUponSubmission")
sAutoResponseWebmaster	= rs("sAutoResponseWebmaster")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sButton) then
check=false
message.AddError("err_mandatory")
end if
if isleeg(sQAalign) then
check=false
message.AddError("err_mandatory")
end if
if (isLeeg(sRedirect) and isLeeg(sFeedback)) or (not isLeeg(sRedirect) and not isLeeg(sFeedback)) then
check=false
message.AddError("err_reg_or_fb")
end if
if bAutoResponder then
if isLeeg(sAutoResponse) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sAutoResponseSubject) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sAutoResponseFromName) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sAutoResponseFromEmail) then
check=false
message.AddError("err_mandatory")
else
if not CheckEmailSyntax(sAutoResponseFromEmail) then
check=false
message.AddError("err_email")
end if
end if
end if
if not isLeeg(sRedirect) then
sRedirect=replace(sRedirect,"http://","",1,-1,1)
sRedirect=replace(sRedirect,"https://","",1,-1,1)
end if
if bSendEmail then
if isLeeg(sTo) then
check=false
message.AddError("err_mandatory")
else
dim arrEmails,emailI
arrEmails=split(sTo,vbcrlf)
for emailI=lbound(arrEmails) to ubound(arrEmails)
if not isLeeg(arrEmails(emailI)) then
if not CheckEmailSyntax(trim(arrEmails(emailI))) then
check=false
message.AddError("err_email")
end if
end if
next
end if
if isLeeg(sSubject) then
check=false
message.AddError("err_mandatory")
end if
end if
End Function
Public Function Save
if isNull(overruleCID) then
if check() then
save=true
else
save=false
exit function
end if
end if
'set db=nothing
'set db=new cls_database
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblForm where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblForm where iId="& iId
end if
rs("sName")	= left(sName,255)
rs("sComment")	= sComment
rs("sCode")	= sCode
rs("sIntro")	= sIntro
rs("sFeedback")	= sFeedback
rs("sTo")	= sTo
rs("sSubject")	= sSubject
rs("sRedirect")	= sRedirect
rs("sRedirectPrefix")	= sRedirectPrefix
rs("sButton")	= sButton
rs("sReset")	= sReset
if not isNull(overruleCID) then
rs("iCustomerID")	= overruleCID
else
rs("iCustomerID")	= cId
end if
rs("dUpdatedTS")	= now()
rs("bSendEmail")	= bSendEmail
rs("bCookie")	= bCookie
rs("bAttachFiles")	= bAttachFiles
rs("bCaptcha")	= bCaptcha
rs("sQAalign")	= sQAalign
rs("sAutoResponse")	= sAutoResponse
rs("bAutoResponder")	= bAutoResponder
rs("sAutoResponseSubject")	= sAutoResponseSubject
rs("sAutoResponseFromName")	= sAutoResponseFromName
rs("sAutoResponseFromEmail")	= sAutoResponseFromEmail
rs("sScriptUponSubmission")	= sScriptUponSubmission
rs("sAutoResponseWebmaster")	= sAutoResponseWebmaster
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function getRequestValues()
sName	= convertStr(Request.Form ("sName"))
sComment	= convertStr(Request.Form ("sComment"))
sCode	= convertStr(Request.Form ("sCode"))
sIntro	= convertStr(Request.Form ("sIntro"))
sFeedback	= convertStr(Request.Form ("sFeedback"))
sTo	= convertStr(Request.Form ("sTo"))
sSubject	= convertStr(Request.Form ("sSubject"))
sRedirect	= convertStr(Request.Form ("sRedirect"))
sRedirectPrefix	= convertStr(Request.Form ("sRedirectPrefix"))
sButton	= convertStr(Request.Form ("sButton"))
sReset	= convertStr(Request.Form ("sReset"))
bSendEmail	= convertBool(Request.Form ("bSendEmail"))
bCookie	= convertBool(Request.Form ("bCookie"))
bAttachFiles	= convertBool(Request.Form("bAttachFiles"))
bCaptcha	= convertBool(Request.Form("bCaptcha"))
sQAalign	= convertStr(Request.Form ("sQAalign"))
bAutoResponder	= convertBool(Request.Form("bAutoResponder"))
sAutoResponse	= convertStr(Request.Form ("sAutoResponse"))
sAutoResponseSubject	= convertStr(Request.Form ("sAutoResponseSubject"))
sAutoResponseFromName	= convertStr(Request.Form ("sAutoResponseFromName"))
sAutoResponseFromEmail	= convertStr(Request.Form ("sAutoResponseFromEmail"))
sScriptUponSubmission	= convertStr(Request.Form ("sScriptUponSubmission"))
sAutoResponseWebmaster	= convertStr(Request.Form ("sAutoResponseWebmaster"))
end function
public function remove
if not isLeeg(iId) then
'submissions verwijderen
dim cSubmissions, key
set cSubmissions=submissions
'velden	verwijderen
dim cFields
set cFields=fields
for each key in cSubmissions
cSubmissions(key).remove(cFields)
next
for each key in cFields
cFields(key).remove
next
dim rs
set rs=db.execute("update tblPage set iFormID=null where iFormID="& iId)
set rs=nothing
set rs=db.execute("update tblCatalog set iFormID=null where iFormID="& iId)
set rs=nothing
set rs=db.execute("delete from tblForm where iId="& iId)
set rs=nothing
end if
end function
public function fields
set fields=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, field, sql
sql="select iId from tblFormField where iFormID=" & iId
sql=sql&" order by iRang asc"
set rs=db.execute(sql)
while not rs.eof
set field=new cls_formField
field.pick(rs(0))
fields.Add field.iID,field
set field=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public function build(action,align,buttonType,itemID)
'On Error Resume next
dim cFields, fFieldKey, postback, showFeedback
postback=false
showFeedback=false
'temporary folder aanmaken om gegevens in te bewaren
dim fso, tempFolder
set fso=server.CreateObject ("scripting.filesystemobject")
tempFolder=GeneratePassWord()
fso.CreateFolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & tempFolder) 
set fso=nothing 
'upload object aanmaken
Dim uploadsDirVar
uploadsDirVar = server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & tempFolder) 
dim Upload
Set Upload = New FreeASPUpload
Upload.Save uploadsDirVar
'exit function
'Response.Write "postback: " & upload.form("postback")
if convertBool(upload.form("postback")) then
if not isLeeg(Upload.form("your message")) then
cleanupASP()
Response.End 
end if
checkCSRF_Upload(Upload.form("QSSEC"))
'form werd gesubmit!
postback=true
'form ophalen
pick(decrypt(upload.form("iFormID")))
if isLeeg(iId) then Response.End 
'check form!
set cFields=fields
'files ophalen
dim UploadedFiles
set UploadedFiles=Upload.UploadedFiles
if bCaptcha then
dim SessionCAPTCHA
SessionCAPTCHA = Trim(Session("CAPTCHA"))
Session("CAPTCHA") = vbNullString
if Len(SessionCAPTCHA) < 1 then
message.AddError("err_captcha")
end if
if lcase(convertStr(SessionCAPTCHA)) <> lcase(convertStr(upload.Form("captcha"))) then
message.AddError("err_captcha")
end if
end if
for each fFieldKey in cFields
if bCookie then
'put in cookie
Response.cookies(encrypt(fFieldKey))=upload.Form (encrypt(fFieldKey))
Response.Cookies(encrypt(fFieldKey)).expires=dateAdd("d",365,date())
end if
'zijn de verplichte ingevuld?
if cFields(fFieldKey).bMandatory and cFields(fFieldKey).sType<>sb_ff_file and cFields(fFieldKey).sType<>sb_ff_image then
if isLeeg(removeEmptyP(upload.Form (encrypt(fFieldKey)))) then
message.AddError("err_mandatory")
end if
end if
'werd een correct emailadres opgegeven?
if cFields(fFieldKey).sType=sb_ff_email then
if not isLeeg(upload.Form (encrypt(fFieldKey))) then
if not CheckEmailSyntax(upload.Form (encrypt(fFieldKey))) then
message.AddError("err_email")
end if
end if
end if
'check de bestanden!
if cFields(fFieldKey).sType=sb_ff_file then
if UploadedFiles.exists(lcase(encrypt(fFieldKey))) then
dim file
set file=UploadedFiles(lcase(encrypt(fFieldKey)))
'check filesize	'
if convertGetal(file.length)>convertGetal(cFields(fFieldKey).iMaxFileSize*1010) then
message.AddError("err_fileSize")
end if
'check empty file
if convertGetal(file.length)=0 then
message.AddError("err_mandatory")
end if
'check type
if not allowedFileTypes.exists(lcase(GetFileExtension(file.FileName))) then
message.AddError("err_fileType")
end if
if not isLeeg(cFields(fFieldKey).sAllowedExtensions) then
if instr(cFields(fFieldKey).sAllowedExtensions,lcase(GetFileExtension(file.FileName)))=0 then
message.AddError("err_fileType")
end if
end if
if message.haserrors then 
file.delete()
end if
elseif cFields(fFieldKey).bMandatory then
'geen file ingevoegd hoewel verplicht!
message.AddError("err_mandatory")
end if
end if
'check de foto's!
if cFields(fFieldKey).sType=sb_ff_image then
if UploadedFiles.exists(lcase(encrypt(fFieldKey))) then
dim image
set image=UploadedFiles(lcase(encrypt(fFieldKey)))
'check filesize	'
if convertGetal(image.length)>convertGetal(cFields(fFieldKey).iMaxFileSize*1010) then '10% margin
message.AddError("err_fileSize")
end if
'check empty file
if convertGetal(image.length)=0 then
message.AddError("err_mandatory")
end if
'check type
select case lcase(GetFileExtension(image.FileName))
case "jpg", "gif", "jpeg"
case else
message.AddError("err_JPGGIF_file")
end select
if message.haserrors then 
image.delete()
end if
elseif cFields(fFieldKey).bMandatory then
'geen image ingevoegd hoewel verplicht!
message.AddError("err_mandatory")
end if
end if
next
if not message.hasErrors then
'submission aanmaken en saven! met die id de velden opvullen
dim submission, theMail
set submission=new cls_submission
submission.iFormID=iId
submission.iItemID=convertLng(itemID)
submission.save()
'gegevens opslaan
dim rs, ts
set rs = db.GetDynamicRS
rs.Open "select * from tblFormFieldValue where 1=2"
ts=now()
for each fFieldKey in cFields
rs.AddNew
rs("iFormFieldId")	= fFieldKey
rs("iSubmissionId")	= submission.iId
select case cFields(fFieldKey).sType
case sb_ff_image,sb_ff_file
if UploadedFiles.exists(lcase(encrypt(fFieldKey))) then
dim saveFile, name
set saveFile=UploadedFiles(lcase(encrypt(fFieldKey)))
name=encrypt(submission.iId) & "_" & encrypt(fFieldKey) & "_" & Generatepassword & "." & GetFileExtension(saveFile.FileName)
'file hernoemen
saveFile.rename name,uploadsDirVar
'file verplaatsen
dim newPath
if isLeeg(cFields(fFieldKey).sFileLocation) then
newPath=replace(uploadsDirVar,tempFolder,"")
newPath=replace(newPath,"\\","\")
if left(newPath,1)="\" and left(newPath,2)<>"\\" then
newPath="\" & newPath
end if
else
newPath=cFields(fFieldKey).sFileLocation & "\"
end if
saveFile.move name,uploadsDirVar,newPath
set saveFile=nothing
'pad bewaren
rs("sValue")=name
end if
case sb_ff_comment 'do nothing
case else
if convertGetal(cFields(fFieldKey).iMaxlength)<>0 then
rs("sValue") = left(convertStr(sanitize(upload.Form (encrypt(fFieldKey)))),convertGetal(cFields(fFieldKey).iMaxlength))
else
rs("sValue") = replace(convertStr(sanitize(upload.Form (encrypt(fFieldKey)))),"_QSDELIMITER","",1,-1,1)
end if
end select
rs.Update
next
rs.close
set rs=nothing
'applicatie runnen?
if customer.bApplication then
if not isLeeg(sScriptUponSubmission) then
dim prepValueForm
for each fFieldKey in cFields
prepValueForm=""
prepValueForm=replace(upload.Form (encrypt(fFieldKey)),"""","'",1,-1,1)
prepValueForm=replace(prepValueForm,vbcrlf," ",1,-1,1)
prepValueForm=replace(prepValueForm,"_QSDELIMITER","",1,-1,1)
sScriptUponSubmission	= Replace(sScriptUponSubmission, "[QS_FORM:" & cFields(fFieldKey).sName &"]",sanitize(prepValueForm),1,-1,1)
next
sScriptUponSubmission	= Replace(sScriptUponSubmission, "[QS_FORM:SUBMISSIONID]",submission.iId,1,-1,1)
call run(sScriptUponSubmission)
end if
end if
if bSendEmail then
'deze volgorde bewaren!
sSubject=insertSubmissionID(sSubject,submission.iId)
sAutoResponseWebmaster=insertSubmissionID(sAutoResponseWebmaster,submission.iId)
sAutoResponseWebmaster=LinkURLs(sAutoResponseWebmaster)
sAutoResponseWebmaster=treatConstants(sAutoResponseWebmaster,true)
'attachments
dim attachments
set attachments=server.createobject("scripting.dictionary")
'mail verzenden!
dim arrEmails,eKey,body
arrEmails=split(sTo,vbcrlf)
'id verzending
body="ID: " & submission.iId & "<br />"
'datum verzending
body=body & l("date") & ": " & formatTimeStamp(now()) & "<br /><br />"
'item ophalen
if isNumeriek(submission.iItemID) then
body=body & l("catalog") & ": " & submission.item.catalog.sName & "<br />"
body=body & submission.item.catalog.sItemName & ": " & submission.item.sTitle & "<br /><br />"
end if
'waarden ophalen
dim cValues, fileURL
set cValues=submission.values(cFields)
'body opvullen in loop!
for each fFieldKey in cFields
sSubject	= Replace(sSubject, "[QS_FORM:" & cFields(fFieldKey).sName &"]",replace(sanitize(upload.Form (encrypt(fFieldKey))),"_QSDELIMITER","",1,-1,1),1,-1,1)
if cFields(fFieldKey).sType<>sb_ff_comment then
fileURL=""
body=body & cleanUpStr(cFields(fFieldKey).sName) & ": "
select case cFields(fFieldKey).sType
case sb_ff_image,sb_ff_file
if not isLeeg(cValues(fFieldKey)) then
if isLeeg(cFields(fFieldKey).sFileLocation) then
fileURL=customer.sVDUrl & application("QS_CMS_userfiles") & cValues(fFieldKey) 
body=body & "<a href='" & fileURL & "'>" & fileURL & "</a>" & "<br />"
attachments.Add generatePassword,server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & cValues(fFieldKey))
else
body=body & "(attached)" & "<br />"
attachments.Add generatePassword, cFields(fFieldKey).sFileLocation & "\" & cValues(fFieldKey)
end if
else
body=body & "<br />"
end if
case sb_ff_checkbox
body=body & convertCheckedYesNo(cValues(fFieldKey)) & "<br />"
'case sb_ff_richtext
'	body=body & removeEmptyP(cValues(fFieldKey)) & "<br />"
case else
body=body & linkUrls(sanitize(cValues(fFieldKey))) & "<br />"
end select
end if
if convertBool(cFields(fFieldKey).bUseForSending) then
fromAddress=cValues(fFieldKey)
end if
next
for eKey=lbound(arrEmails) to ubound(arrEmails)
if not isLeeg(trim(arrEmails(eKey))) then
if CheckEmailSyntax(trim(arrEmails(eKey))) then
set theMail=new cls_mail_message
if bAttachFiles then
set theMail.attachments=attachments
end if
if not isLeeg(fromAddress) then 
theMail.fromemail=fromAddress
theMail.fromname=fromAddress
end if
theMail.receiver=trim(arrEmails(eKey))
theMail.subject=sSubject
theMail.body=sAutoResponseWebmaster & "<br />" & body & "<hr /><p><b><u>Visitor &amp; Server Data:</u></b></p>" & submission.sVisitorDetails
theMail.send()
set theMail=nothing
end if
end if
next
set attachments=nothing
end if
if bAutoResponder then
sAutoResponseSubject=insertSubmissionID(sAutoResponseSubject,submission.iId)
sAutoResponse=insertSubmissionID(sAutoResponse,submission.iId)
sAutoResponse=LinkURLs(sAutoResponse)
sAutoResponse=treatConstants(sAutoResponse,true)
for each fFieldKey in cFields
sAutoResponseSubject	= Replace(sAutoResponseSubject, "[QS_FORM:" & cFields(fFieldKey).sName &"]",sanitize(replace(upload.Form (encrypt(fFieldKey)),"_QSDELIMITER","",1,-1,1)),1,-1,1)
sAutoResponse	= Replace(sAutoResponse, "[QS_FORM:" & cFields(fFieldKey).sName &"]",sanitize(replace(upload.Form (encrypt(fFieldKey)),"_QSDELIMITER","",1,-1,1)),1,-1,1)
next

if instr(sAutoResponse,"[QS_COPYSUBMISSION]")<>0 then
	sAutoResponse	= Replace(sAutoResponse, "[QS_COPYSUBMISSION]",submission.sCopyValues,1,-1,1)
end if

for each fFieldKey in cFields
if cFields(fFieldKey).bAutoResponder then
if not isLeeg(convertStr(upload.Form (encrypt(fFieldKey)))) then
'send autoresponse
set theMail=new cls_mail_message
theMail.fromname	= sAutoResponseFromName
theMail.fromemail	= sAutoResponseFromEmail
theMail.receiver	= convertStr(upload.Form (encrypt(fFieldKey)))
theMail.subject	= sAutoResponseSubject
theMail.body	= sAutoResponse
theMail.send()
set theMail=nothing
end if
end if
next
end if
'temp folder verwijderen
if not deleteTempFolder(uploadsDirVar)	then
ErrorReport "deleteTempFolder",err
end if
'opruimen
'set submission=nothing
'feedback message tonen of redirecten naar verwijzigingspagina
if not isLeeg(sFeedback) then
showFeedback=true
else
Response.Redirect (sRedirectPrefix & sRedirect)
end if
end if
else
set cFields=fields
end if
'anker plaatsen
build="<a name='form"& encrypt(iId) &"'></a>" 
if showFeedback then
build=build&"<div id='QS_form'>"

sFeedback=Replace(sFeedback, "[QS_COPYSUBMISSION]",submission.sCopyValues,1,-1,1)

build=build&"<p>" & show(sFeedback) & "</p>"
build=build&"</div>"
else
'temp folder verwijderen
if not deleteTempFolder(uploadsDirVar)	then
ErrorReport "deleteTempFolder",err
end if
build=build&"<div id='QS_form'>"
build=build&"<form enctype=" & """" & "multipart/form-data" & """" & " action="& """" & action & """" &" method='post' name='quickerform'>"
build=build&QS_secCodeHidden
build=build&"<input type='hidden' name='iFormID' value=" & """" & encrypt(iId) & """" & " />"
build=build&"<input type='hidden' name='formAction' value=" & """" & sButton & """" & " />"
build=build&"<input type='hidden' name='postBack' value='"&true&"' />"
build=build&"<input style='display:none' type='text' name='your message' value='' />"
build=build&"<p>" & show(sIntro) & "</p>"
dim defaultValue, JSAdded
JSAdded=false
for each fFieldKey in cFields
if cFields(fFieldKey).sType<>sb_ff_hidden then
build=build & "<div class='QS_fieldline'>"
build=build & "<div "
if sQAalign=QS_QtopAbottom or cFields(fFieldKey).sType=sb_ff_comment then
build=build & "class='QS_oneline'"
else
build=build & "class='QS_fieldlabel'"
end if
build=build&">" 
if cFields(fFieldKey).sType=sb_ff_comment then
if  isLeeg(cFields(fFieldKey).sValues) then
build=build&cFields(fFieldKey).sName 'backwards compatibilty
else
build=build&cFields(fFieldKey).sValues
end if
else
build=build&cFields(fFieldKey).sName
end if
if cFields(fFieldKey).bMandatory then
build=build&"*"
end if
if cFields(fFieldKey).sType=sb_ff_file or cFields(fFieldKey).sType=sb_ff_image then
build=build&"&nbsp;<span style=""font-size:0.8em"">(max. " & cFields(fFieldKey).iMaxFileSize &" kB)</span>"
end if
build=build & "</div>"
if cFields(fFieldKey).sType<>sb_ff_comment and cFields(fFieldKey).sType<>sb_ff_hidden then
build=build & "<div class='QS_fieldvalue'>"
if postback then
defaultValue=upload.form(encrypt(fFieldKey))
else
defaultValue=Request.Cookies (encrypt(fFieldKey))
end if
if not isLeeg(cFields(fFieldKey).sPlaceHolder) then
cFields(fFieldKey).sPlaceHolder=treatConstants(cFields(fFieldKey).sPlaceHolder,true)
end if
select case cFields(fFieldKey).sType

case sb_ff_text, sb_ff_url, sb_ff_email
	
	build=build&"<input "
	
	if cFields(fFieldKey).bMandatory then
		build=build & " required "
	end if
	
	if not isLeeg(cFields(fFieldKey).sPlaceHolder) then
		build=build&" placeholder=""" & sanitize(cFields(fFieldKey).sPlaceHolder) & """ "
	end if
	
	build=build& " type='"
	
	select case cFields(fFieldKey).sType
	
		case sb_ff_text
			build=build & "text"
		
		case sb_ff_email
			build=build & "email"
		
		case sb_ff_url
			build=build & "url"
		
	end select
	
	build=build& "' "	
	build=build& " size='"& cFields(fFieldKey).iSize &"' maxlength='"& cFields(fFieldKey).iMaxLength &"' name=" & """" & encrypt(fFieldKey) & """" &" value=" & """" & quotRep(defaultValue) & """" & " />"
	resetJS=resetJS&"document.quickerform." & encrypt(fFieldKey) & ".value='';"

case sb_ff_textarea,sb_ff_richtext

	build=build& "<textarea "

	if cFields(fFieldKey).bMandatory then
		build=build & " required "
	end if
	
	if not isLeeg(cFields(fFieldKey).sPlaceHolder) then
		build=build&" placeholder=""" & sanitize(cFields(fFieldKey).sPlaceHolder) & """ "
	end if

	'show maximum-left box
	if convertGetal(cFields(fFieldKey).iMaxLength)<>0 then
		build=build&" onKeyDown=""javascript:textCounter(this.form." & encrypt(fFieldKey) & ",'" & encrypt(fFieldKey) & "remLen',"&cFields(fFieldKey).iMaxLength&");"" onKeyUp=""javascript:textCounter(this.form." & encrypt(fFieldKey) & ",'" & encrypt(fFieldKey) & "remLen',"&cFields(fFieldKey).iMaxLength&");"" cols='"& cFields(fFieldKey).iCols &"' rows='"& cFields(fFieldKey).iRows &"' name='" & encrypt(fFieldKey) &"'>" & quotRep(defaultValue) &"</textarea>"
		build=build&"<br /><span id='" & encrypt(fFieldKey) & "remLen'>" & cFields(fFieldKey).iMaxLength & "</span>&nbsp;" & l("charactersleft") 
	else
		build=build&" cols='"& cFields(fFieldKey).iCols &"' rows='"& cFields(fFieldKey).iRows &"' name='" & encrypt(fFieldKey) &"'>" & quotRep(defaultValue) &"</textarea>"
	end if
	
	resetJS=resetJS&"document.quickerform." & encrypt(fFieldKey) & ".value='';"

case sb_ff_checkbox

	build=build&"<input "
	
	if cFields(fFieldKey).bMandatory then
		build=build & " required "
	end if	
	
	build=build & " type='checkbox' name="& """" & encrypt(fFieldKey) &""""&" value='checked' "& defaultValue &" />"
	resetJS=resetJS&"document.quickerform." & encrypt(fFieldKey) & ".checked=false;"

case sb_ff_select

	build=build&"<select "
	
	if cFields(fFieldKey).bMandatory then
		build=build & " required "
	end if
	
	
	build=build& " name="& """" & encrypt(fFieldKey) &"""" &">"& cFields(fFieldKey).showSelected(defaultValue) &"</select>"
	resetJS=resetJS&"document.quickerform." & encrypt(fFieldKey) & ".selectedIndex=0;"

case sb_ff_radio

	if convertBool(cFields(fFieldKey).bAllowMS) then
		build=build& cFields(fFieldKey).showSelectedRadioCB(defaultValue)
	else
		build=build& cFields(fFieldKey).showSelectedRadio(defaultValue)
	end if
	
case sb_ff_date

	build=build&"<input "
	
	if cFields(fFieldKey).bMandatory then
		build=build & " required "
	end if
	
	build=build&" type=""text"" id=""" & encrypt(fFieldKey) & """ name=""" & encrypt(fFieldKey) & """ value=""" & defaultValue & """ />"& JQDatePicker(encrypt(fFieldKey))
	resetJS=resetJS&"document.quickerform." & encrypt(fFieldKey) & ".value='';"

case sb_ff_file,sb_ff_image

	build=build&"<input "
	
	if cFields(fFieldKey).bMandatory then
		build=build & " required "
	end if
	
	build=build&" type='file' name=" & """" & encrypt(fFieldKey) & """" &" />"
	
	if not isLeeg(cFields(fFieldKey).sAllowedExtensions) then
		build=build & "<br /><span style=""font-size:0.8em"">" & l("allowedfiletypes") & ": " & cFields(fFieldKey).sAllowedExtensions & "</span>"
	end if
	
	resetJS=resetJS&"document.quickerform." & encrypt(fFieldKey) & ".value='';"
	
end select
build=build & "</div>"
end if
build=build & "</div>"
else
build=build & "<input type=""hidden"" name=""" & encrypt(fFieldKey) & """ value=""" & sanitize(cFields(fFieldKey).sValues) & """ />"
end if
next
if bCaptcha then
build=build & "<div class='QS_fieldline'>"
build=build&"<div class='"
if sQAalign=QS_QtopAbottom then
build=build & "QS_oneline"
else
build=build & "QS_fieldlabel"
end if
build=build & "'>"& l("captcha") &":*</div>"
build=build & "<div class='QS_fieldvalue'><img src='" & C_DIRECTORY_QUICKERSITE & "/asp/includes/captcha.asp' style='vertical-align: middle;margin:0px;border-style:none' alt='This is a CAPTCHA Image' />&nbsp;<input value="""& sanitize(upload.form("captcha")) &""" type='text' size='5' style='width:63px' maxlength='4' name='captcha' required /></div>"
build=build & "</div>"
end if
build=build & "<div class='QS_fieldline'>"
build=build&"<div class='"
if sQAalign=QS_QtopAbottom then
build=build & "QS_oneline"
else
build=build & "QS_fieldlabel"
end if
build=build&"'>&nbsp;</div>"
build=build&"<div class='QS_fieldvalue'><span id=""mandSPAN" & iId & """ style=""font-size:0.85em"">(*) "& l("mandatory") &"</span></div>"
build=build & "</div>"
build=build & "<div class='QS_fieldline'>"
build=build&"<div class='"
if sQAalign=QS_QtopAbottom then
build=build & "QS_oneline"
else
build=build & "QS_fieldlabel"
end if
build=build&"'>&nbsp;</div>"
build=build&"<div class='QS_fieldvalue'><input class='art-button' type='"& buttonType &"' value=" & """" & quotRep(sButton) & """" &" name='dummy' />"
if not isLeeg(sReset) then
build=build&"&nbsp;<input type='button' class='art-button' onclick='javascript:resetForm();' value=" & """" & quotRep(sReset) & """" &" name='reset' />"
end if
build=build&"</div>"
build=build&"</div>"
build=build&"</form></div>"
build=build&"<script type='text/javascript'>"
build=build&"function resetForm() {"
build=build&resetJS
build=build&"}"
build=build&"</script>"
end if
Set Upload=nothing
'delete all tempfolders
dim fsoDelete,subDELfolder,DELfolder
set fsoDelete=server.CreateObject ("scripting.filesystemobject")
if fsoDelete.folderexists(server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles"))) then
set DELfolder=fsoDelete.getfolder(server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles")))
for each subDELfolder in DELfolder.subfolders
if len(subDELfolder.name)=8  then
if subDELfolder.size=0 then
subDELfolder.delete()
end if 
end if
next
set DELfolder=nothing
end if
set fsoDelete=nothing
On Error Goto 0
end function
private function deleteTempFolder(stempfolder)
deleteTempFolder=true
dim fsoDelete
set fsoDelete=server.CreateObject ("scripting.filesystemobject")
if fsoDelete.FolderExists (stempFolder) then
on error resume next
if right(stempFolder,1)="\" then
fsoDelete.Deletefolder left(stempFolder,len(stempFolder)-1)
else
fsoDelete.Deletefolder stempFolder
end if
if err.number=0 then
deleteTempFolder=true
end if
on error goto 0
end if
set fsoDelete=nothing
end function
public function submissions
set submissions=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, submission
set rs=db.execute("select iId from tblFormSubmission where iFormID="& iId & " order by dCreatedTS desc")
while not rs.eof
set submission=new cls_submission
submission.pick(rs(0))
submissions.Add submission.iId, submission
set submission=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public function copy()
if isNumeriek(iId) then
dim cFields, fKey
set cFields=fields
iId=null
sName=l("copyof") & " " & sName
save()
for each fKey in cFields
cFields(fKey).copy(iId)
next
set cFields=nothing
end if
end function
public function copyToCustomer(oCID)
if isNumeriek(iId) then
dim oldId
oldID=iId
overruleCID	= oCID
dim cFields, fKey
set cFields=fields
iId=null
save()
for each fKey in cFields
cFields(fKey).copy(iId)
next
set cFields=nothing
db.execute("update tblPage set iFormID=" & iId & " where iFormID=" & oldID & " and iCustomerID=" & oCID)
db.execute("update tblCatalog set iFormID=" & iId & " where iFormID=" & oldID & " and iCustomerID=" & oCID)
end if
end function
public function insertSubmissionID(sResponse,iSubmissionID)
insertSubmissionID=Replace(convertStr(sResponse), "[QS_FORM:SUBMISSIONID]",convertStr(iSubmissionID),1,-1,1)
end function
private sub run(script)
execute(treatConstants(script,true))
end sub
public function removeAllSubmissions
dim cSubmissions, key
set cSubmissions=submissions
'velden	verwijderen
dim cFields
set cFields=fields
for each key in cSubmissions
cSubmissions(key).remove(cFields)
next
set cSubmissions=nothing
set cFields=nothing
end function
end class%>
