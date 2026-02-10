
<%class cls_contact
Public iId, iCustomerID,sEmail,sOrigEmail,sNickName,sPw,dUpdatedTS,dCreatedTS,dLastLoginTS,fields,iStatus,allowDE,copyStatus, bGetEmailsFromSite, dLogoutTS, sAnName, sAvatar
private p_bHasSubscribedToTheme
Private Sub Class_Initialize
on error resume next
p_bHasSubscribedToTheme=null
iId=null
set fields=server.CreateObject ("scripting.dictionary")
iCustomerID=cId
iStatus=customer.iDefaultStatus
allowDE=false
dLastLoginTS=null
dLogoutTS=null
bGetEmailsFromSite=false
sAvatar=""
pick(decrypt(request("iContactID")))
on error goto 0
end sub
public function quickPick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblContact where iCustomerID="&cid&" and iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
iCustomerID	= rs("iCustomerID")
sEmail	= rs("sEmail")
sNickName	= rs("sNickName")
sOrigEmail	= rs("sOrigEmail")
sPw	= rs("sPw")
dUpdatedTS	= rs("dUpdatedTS")
dCreatedTS	= rs("dCreatedTS")
dLastLoginTS	= rs("dLastLoginTS")
dLogoutTS	= rs("dLogoutTS")
iStatus	= rs("iStatus")
bGetEmailsFromSite	= rs("bGetEmailsFromSite")
sAvatar	= rs("sAvatar")
copyStatus	= iStatus
if iCustomerID<>cId then Response.End 
end if
set RS = nothing
end if
end function
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblContact where iCustomerID="&cid&" and iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
iCustomerID	= rs("iCustomerID")
sEmail	= rs("sEmail")
sNickName	= rs("sNickName")
sOrigEmail	= rs("sOrigEmail")
sPw	= rs("sPw")
dUpdatedTS	= rs("dUpdatedTS")
dCreatedTS	= rs("dCreatedTS")
dLastLoginTS	= rs("dLastLoginTS")
dLogoutTS	= rs("dLogoutTS")
iStatus	= rs("iStatus")
bGetEmailsFromSite	= rs("bGetEmailsFromSite")
sAvatar	= rs("sAvatar")
copyStatus	= iStatus
if iCustomerID<>cId then Response.End 
dim rsFields, sValue,iFieldID
set rsFields=db.execute("select iFieldId,sValue from tblContactValues where iContactID="&iId)
fields.RemoveAll()
while not rsFields.eof
iFieldID	= rsFields(0)
sValue	= rsFields(1)
fields.Add iFieldID,sValue
rsFields.movenext
wend
set rsFields=nothing
end if
set RS = nothing
end if
end function
Public Function Check(contactFields)
Check = true
if isLeeg(iCustomerID) then
check=false
end if
dim contactField
for each contactField in contactFields
if contactFields(contactField).bMandatory then
if isLeeg(fields(contactField)) then
message.AddError("err_mandatory")
check=false
end if
end if
next
if isLeeg(sEmail) then
check=false
message.AddError("err_mandatory")
elseif not CheckEmailSyntax(sEmail) then
message.AddError("err_email")
check=false
else
sEmail=lcase(sEmail)
end if
if isLeeg(sPw) then
message.AddError("err_mandatory")
check=false
end if
if isLeeg(sNickName) then
message.AddError("err_mandatory")
check=false
end if
if check then
if not allowDE then
'check double email
dim rs
set rs=db.execute("select iId from tblContact where iId<>"&convertGetal(iId)&" and iCustomerID="& cId &" and sEmail='" & sEmail & "'")
if not rs.eof then
message.AddError("err_doubleemail")
check=false
end if
end if
'check double nickname
set rs=db.execute("select iId from tblContact where iId<>"&convertGetal(iId)&" and iCustomerID="& cId &" and sNickName='" & cleanup(sNickName) & "'")
if not rs.eof then
message.AddError("err_doublenickname")
check=false
end if
end if
End Function
Public function lastLoginTSSave()
on error resume next
if convertGetal(iId)<>0 then
dim rs
set rs = db.GetDynamicRS
rs.Open "select dLastLoginTS from tblContact where iId="& iId
rs(0) = now()
rs.Update 
set rs=nothing
end if
on error goto 0
end function
Public function lastLogoutTSSave()
on error resume next
if convertGetal(iId)<>0 then
dim rs
set rs = db.GetDynamicRS
rs.Open "select dLogoutTS from tblContact where iId="& iId
rs(0) = now()
rs.Update 
set rs=nothing
end if
on error goto 0
end function
Public Function Save(contactFields)
if check(contactFields) then
save=true
else
save=false
exit function
end if
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblContact where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblContact where iId="& iId
end if
rs("sEmail")	= sEmail
rs("bGetEmailsFromSite")	= bGetEmailsFromSite
rs("sNickName")	= sNickName
rs("sOrigEmail")	= sOrigEmail
rs("sPw")	= sPw
rs("iCustomerID")	= cId
rs("dUpdatedTS")	= now()
rs("iStatus")	= iStatus
rs("dLastLoginTS")	= dLastLoginTS
dUpdatedTS=rs("dUpdatedTS")
dCreatedTS=rs("dCreatedTS")
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
'fields
dim rsFields
set rsFields=db.execute("delete from tblContactValues where iContactID="&iId)
set rsFields=nothing
set rsFields=db.GetDynamicRS
rsFields.Open "select * from tblContactValues where 1=2"
dim field
for each field in fields
rsFields.AddNew
rsFields("iContactID")=iID
rsFields("iFieldID")=field
rsFields("sValue")=fields(field)
rsFields.Update
next
rsFields.close
set rsFields=nothing
'check double email
set rs=db.execute("select iId from tblContact where iId<>"&convertGetal(iId)&" and iCustomerID=" & cId & " and sEmail='"& sEmail &"'")
if not rs.eof then
dim removeContact
set removeContact=new cls_contact
removeContact.Pick(rs(0))
removeContact.delete()
set removeContact=nothing
end if
set rs=nothing
'check email?
if copyStatus<>iStatus and iStatus<>cs_silent then
dim iMess
set iMess=new cls_customerImess
iMess.pick(iStatus)
if iMess.bEnabled and not isLeeg(iMess.sBody) then
dim wMail
set wMail=new cls_mail_message
wMail.receiver=sEmail
wMail.subject=iMess.sSubject
wMail.body=iMess.sBody
wMail.send
set wMail=nothing
end if
end if
end function
public function getRequestValues(byref contactFields)
dim contactField
for each contactField in contactFields
if not convertBool(contactFields(contactField).bProfile) and not saveHiddenValues then
if not saveHiddenValues and convertGetal(iId)=0 then
fields(contactField)=""
end if
else
select case contactFields(contactField).sType
case sb_date
fields(contactField)=convertCalcDate(Request.Form(encrypt(contactField)))
case else
fields(contactField)=Request.Form(encrypt(contactField))
end select
end if
next
sEmail	= trim(convertStr(Request.Form ("sEmail")))
sNickName	= convertStr(Request.Form("sNickName"))
sPw	= convertStr(Request.Form ("sPw"))
bGetEmailsFromSite	= convertBool(Request.Form ("bGetEmailsFromSite"))
end function
public function delete
if isNumeriek(iId) then
removeAvatar()
dim rs
set rs=db.execute("delete from tblContactValues where iContactID="&iId)
set rs=nothing
dim cposts,postKey
set cposts=posts
for each postKey in cposts
cposts(postKey).remove
next
set cposts=nothing
set rs=db.execute("update tblTheme set iContactID=null where iContactID="&iId)
set rs=nothing
set rs=db.execute("delete from tblThemeSubscription where iContactID="&iId)
set rs=nothing
set rs=db.execute("delete from tblThemeTopicSubscription where iContactID="&iId)
set rs=nothing
set rs=db.execute("delete from tblContactPage where iContactID="&iId)
set rs=nothing
set rs=db.execute("delete from tblContact where iId="&iId)
set rs=nothing
set rs=db.execute("delete from tblContactRegistration where iCustomerID="&cId&" and sEmail='"&sEmail&"'")
end if
end function
Public function posts
set posts=server.CreateObject ("scripting.dictionary")
if convertGetal(iId)<>0 then
dim rs,sql,post
sql="select iId from tblPost where iContactID="& iId
set rs=db.execute(sql)
while not rs.eof
set post=new cls_post
post.pick(rs(0))
posts.Add  post.iId, post
set post=nothing
rs.movenext
wend
set rs=nothing
end if
end function
Public sub resetPW
sPw=GeneratePassWord()
dim rs
set rs=db.execute("update tblContact set sPw='"&sPw&"' where iId="&iId)
set rs=nothing
dim theMail, body
set theMail=new cls_mail_message
theMail.receiver=sEmail
theMail.subject=l("password")&" "&customer.siteName 
if instr(customer.intranetPWEmail,"[QS_intranet:contactemail]")<>0 then
body=replace(treatConstants(customer.intranetPWEmail,true),"[QS_intranet:contactemail]",sEmail,1,-1,1)
body=replace(body,"[QS_intranet:contactpassword]",sPw,1,-1,1)
theMail.body=body
else
theMail.body=LinkURLs(treatConstants(customer.intranetPWEmail,true)) & "<br />"& "<br />" & "Email: "& sEmail & "<br />"& "<br />" & l("password")&": "&	sPw
end if
theMail.send()
set theMail=nothing
end sub
public function bHasSubscribedToTheme(iThemeID)
if isNumeriek(iId) then
if isNull(p_bHasSubscribedToTheme) then
dim rs
set rs=db.execute("select count(*) from tblThemeSubscription where iThemeId="& convertGetal(iThemeID) &" and iContactId="& convertGetal(iId))
if clng(rs(0))=0 then 
p_bHasSubscribedToTheme=false
else
p_bHasSubscribedToTheme=true
end if
set rs=nothing
end if
bHasSubscribedToTheme=p_bHasSubscribedToTheme
else
bHasSubscribedToTheme=false
end if
end function
public function bHasSubscribedToTopic(iTopicID)
if isNumeriek(iId) then
dim rs
set rs=db.execute("select count(*) from tblThemeTopicSubscription where iPostID="& convertGetal(iTopicID) &" and iContactId="& convertGetal(iId))
if clng(rs(0))=0 then 
bHasSubscribedToTopic=false
else
bHasSubscribedToTopic=true
end if
set rs=nothing
else
bHasSubscribedToTopic=false
end if
end function
public function getTPer
set getTPer=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs,iPageID
set rs=db.execute("select iTitleID from tblContactPage where iContactID=" & iId)
while not rs.eof
iPageID=convertGetal(rs(0))
if convertGetal(iPageID)<>0 then
getTPer.Add iPageID,""
end if
rs.movenext
wend
set rs=nothing
end if
end function
public function getBPer
set getBPer=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs,iPageID
set rs=db.execute("select iBodyID from tblContactPage where iContactID=" & iId)
while not rs.eof
iPageID=rs(0)
if convertGetal(iPageID)<>0 then
getBPer.Add iPageID,""
end if
rs.movenext
wend
set rs=nothing
end if
end function
public function getLPer
set getLPer=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs,iPageID
set rs=db.execute("select iLPid from tblContactPage where iContactID=" & iId)
while not rs.eof
iPageID=rs(0)
if convertGetal(iPageID)<>0 then
getLPer.Add iPageID,""
end if
rs.movenext
wend
set rs=nothing
end if
end function
public function savePermissions(iBodyID,iTitleID,iLPID)
savePermissions=true
dim rs
set rs=db.execute("delete from tblContactPage where iContactId=" & iId)
set rs=nothing
'save bodyID's
dim el
set rs=db.getDynamicRS
rs.open "select * from tblContactPage where iContactID is null"
for each el in iBodyID
rs.addNew()
rs("iContactID")=iId
rs("iBodyID")=convertGetal(replace(el,"iPageIDBody","",1,-1,1))
rs.update()
next  
rs.close
set rs=nothing
'save titleID's
set rs=db.getDynamicRS
rs.open "select * from tblContactPage where iContactID is null"
for each el in iTitleID
rs.addNew()
rs("iContactID")=iId
rs("iTitleID")=convertGetal(replace(el,"iPageIDTitle","",1,-1,1))
rs.update()
next  
rs.close
set rs=nothing
'save LP'ids
set rs=db.getDynamicRS
rs.open "select * from tblContactPage where iContactID is null"
for each el in iLPID
rs.addNew()
rs("iContactID")=iId
rs("iLPID")=convertGetal(replace(el,"iPageIDLP","",1,-1,1))
rs.update()
next  
rs.close
set rs=nothing
createUserFilesFolder()
if err.number<>0 then
savePermissions=false
end if
end function
public sub createUserFilesFolder
dim fso
set fso=server.CreateObject ("scripting.filesystemobject")
if not fso.FolderExists (server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles")) then
fso.CreateFolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles") 
end if 
if not fso.FolderExists (server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/" & iId)) then
fso.CreateFolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/" & iId) 
end if
set fso=nothing
end sub
public property get sClickNickName
if convertGetal(iId)=0 then
sClickNickName=quotrep(sAnName)
else
if logon.authenticatedIntranet and convertBool(bGetEmailsFromSite) and iId<>logon.contact.iId then
sClickNickName="<a class=""bPopupFullWidthNoReload"" href=""" & C_DIRECTORY_QUICKERSITE & "/asp/fs_mailcontact.asp?iContactID=" & encrypt(iId) & """>" & quotrep(sNickname) & "</a>"
else
sClickNickName=quotrep(sNickName)
end if
end if
end property
public function removeAvatar
dim rs
set rs=db.execute("update tblContact set sAvatar='' where iId=" & iId)
set rs=nothing
'remove file
dim fso
set fso=server.CreateObject ("scripting.filesystemobject")
if fso.FileExists (server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/" & iId & "_" & sAvatar & ".jpg")) then
fso.DeleteFile server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/" & iId & "_" & sAvatar & ".jpg")
end if
if fso.FileExists (server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/" & iId & "_" & sAvatar)) then
fso.DeleteFile server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/" & iId & "_" & sAvatar)
end if
sAvatar=""
set fso=nothing
end function
public function saveAvatar(cpw)
removeAvatar()
dim rs
set rs=db.execute("update tblContact set sAvatar='" & cpw & "' where iId=" & iId)
set rs=nothing
sAvatar=cpw
end function
public property get sImgTagAvatar(size)
if not isLeeg(customer.sAvatarBorderColor) then
sImgTagAvatar= "<img style=""margin:0px;border:1px solid " & customer.sAvatarBorderColor & """ "
else
sImgTagAvatar= "<img style=""margin:0px;border-style:none"" "
end if
select case right(sAvatar,3)
case "gif"
sImgTagAvatar=sImgTagAvatar& " src=""" & C_VIRT_DIR &  Application("QS_CMS_userfiles") & "userfiles/"
sImgTagAvatar=sImgTagAvatar& iId & "_" & sAvatar & """ alt=""" & quotrep(sNickname) & """ width=""" & size & """ height=""" & size & """ style=""height:" & size & "px;width:" & size & "px"" title=""" & quotrep(sNickname) & """ />"
case "jpg"
sImgTagAvatar=sImgTagAvatar& " src=""" & C_DIRECTORY_QUICKERSITE
sImgTagAvatar=sImgTagAvatar& "/showthumb.aspx?FSR=1&amp;maxsize=" & size & "&amp;img=" & C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/"
sImgTagAvatar=sImgTagAvatar& iId & "_" & sAvatar & """ alt=""" & quotrep(sNickname) & """ title=""" & quotrep(sNickname) & """ />"
case else
sImgTagAvatar=sImgTagAvatar& " src=""" & C_DIRECTORY_QUICKERSITE
sImgTagAvatar=sImgTagAvatar& "/showthumb.aspx?FSR=1&amp;maxsize=" & size & "&amp;img=" & C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/"
sImgTagAvatar=sImgTagAvatar& iId & "_" & sAvatar & ".jpg"" alt=""" & quotrep(sNickname) & """ title=""" & quotrep(sNickname) & """ />"
end select
end property
public function getAvatar
if not customer.bUseAvatars then exit function
if not isLeeg(sAvatar) then
select case right(sAvatar,3) 
case "gif"
getAvatar="<a title=""" & quotrep(sNickname) & """ href="""& C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/" & iId & "_" & sAvatar & """ "
getAvatar=getAvatar& "class=""QSPPIMG""><img style=""margin:0px;"
if not isLeeg(customer.sAvatarBorderColor) then
getAvatar=getAvatar& "border:1px solid " & customer.sAvatarBorderColor & """ "
else
getAvatar=getAvatar& "border-style:none"" "
end if
getAvatar=getAvatar& " style=""width:" & customer.iAvatarSize & "px;height:" & customer.iAvatarSize & "px"" height=""" & customer.iAvatarSize & """ width=""" & customer.iAvatarSize & """ src=""" &  C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/"
getAvatar=getAvatar& iId & "_" & sAvatar & """ alt=""" & quotrep(sNickname) & """ />"
getAvatar=getAvatar& "</a>"
case "jpg"
getAvatar="<a title=""" & quotrep(sNickname) & """ href="""& C_DIRECTORY_QUICKERSITE & "/showthumb.aspx?FSR=0&amp;"
getAvatar=getAvatar& "maxsize=600&amp;img=" & C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/" & iId & "_" & sAvatar & """ "
getAvatar=getAvatar& "class=""QSPPIMG""><img style=""margin:0px;"
if not isLeeg(customer.sAvatarBorderColor) then
getAvatar=getAvatar& "border:1px solid " & customer.sAvatarBorderColor & """ "
else
getAvatar=getAvatar& "border-style:none"" "
end if
getAvatar=getAvatar& " src=""" & C_DIRECTORY_QUICKERSITE
getAvatar=getAvatar& "/showthumb.aspx?FSR=1&amp;maxsize=" & customer.iAvatarSize & "&amp;img=" & C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/"
getAvatar=getAvatar& iId & "_" & sAvatar & """ alt=""" & quotrep(sNickname) & """ />"
getAvatar=getAvatar& "</a>"
case else
getAvatar="<a title=""" & quotrep(sNickname) & """ href="""& C_DIRECTORY_QUICKERSITE & "/showthumb.aspx?FSR=0&amp;"
getAvatar=getAvatar& "maxsize=600&amp;img=" & C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/" & iId & "_" & sAvatar & ".jpg"" "
getAvatar=getAvatar& "class=""QSPPIMG""><img style=""margin:0px;"
if not isLeeg(customer.sAvatarBorderColor) then
getAvatar=getAvatar& "border:1px solid " & customer.sAvatarBorderColor & """ "
else
getAvatar=getAvatar& "border-style:none"" "
end if
getAvatar=getAvatar& " src=""" & C_DIRECTORY_QUICKERSITE
getAvatar=getAvatar& "/showthumb.aspx?FSR=1&amp;maxsize=" & customer.iAvatarSize & "&amp;img=" & C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/"
getAvatar=getAvatar& iId & "_" & sAvatar & ".jpg"" alt=""" & quotrep(sNickname) & """ />"
getAvatar=getAvatar& "</a>"
end select
else
getAvatar="<img height=""" & customer.iAvatarsize & """ width=""" & customer.iAvatarsize & """ style=""width:"&customer.iAvatarsize&";height:"&customer.iAvatarsize&";margin:0px;"
if not isLeeg(customer.sAvatarBorderColor) then
getAvatar=getAvatar& "border:1px solid " & customer.sAvatarBorderColor & """ "
else
getAvatar=getAvatar& "border-style:none"" "
end if
dim h
set h=new md5
getAvatar=getAvatar& " src=""http://www.gravatar.com/avatar.php?gravatar_id=" & h.hash(sEmail) & "&amp;default=" & customer.sQSUrl & "/fixedImages/avatar.jpg"" alt=""avatar"" />"
set h=nothing
end if
end function
public function getClickAvatar
if not customer.bUseAvatars then exit function
if not isLeeg(sAvatar) then
getClickAvatar="<a title=""" & quotrep(sNickname) & """ href=""default.asp?pageAction=avataredit"" "
getClickAvatar=getClickAvatar& "class=""QSPPAVATAR"">"
getClickAvatar=getClickAvatar& "<img style=""margin:0px;"
if not isLeeg(customer.sAvatarBorderColor) then
getClickAvatar=getClickAvatar& "border:1px solid " & customer.sAvatarBorderColor & """ "
else
getClickAvatar=getClickAvatar& "border-style:none"" "
end if
select case right(sAvatar,3)
case "gif"
getClickAvatar=getClickAvatar& "src="""
getClickAvatar=getClickAvatar& C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/"
getClickAvatar=getClickAvatar& iId & "_" & sAvatar & """ height=""" & customer.iAvatarSize & """ width=""" & customer.iAvatarSize & """ style=""width:" & customer.iAvatarSize & "px;height:" & customer.iAvatarSize & "px"" alt=""" & quotrep(sNickname) & """ title=""" & quotrep(sNickname) & """ />"
getClickAvatar=getClickAvatar& "</a>"
case "jpg"
getClickAvatar=getClickAvatar& "src=""" & C_DIRECTORY_QUICKERSITE
getClickAvatar=getClickAvatar& "/showthumb.aspx?FSR=1&amp;maxsize=" & customer.iAvatarSize & "&amp;img=" & C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/"
getClickAvatar=getClickAvatar& iId & "_" & sAvatar & """ alt=""" & quotrep(sNickname) & """ title=""" & quotrep(sNickname) & """ />"
getClickAvatar=getClickAvatar& "</a>"
case else
getClickAvatar=getClickAvatar& "src=""" & C_DIRECTORY_QUICKERSITE
getClickAvatar=getClickAvatar& "/showthumb.aspx?FSR=1&amp;maxsize=" & customer.iAvatarSize & "&amp;img=" & C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/"
getClickAvatar=getClickAvatar& iId & "_" & sAvatar & ".jpg"" alt=""" & quotrep(sNickname) & """ title=""" & quotrep(sNickname) & """ />"
getClickAvatar=getClickAvatar& "</a>"
end select
else
getClickAvatar="<a href=""default.asp?pageAction=avataredit"" class=""QSPPAVATAR""><img height=""" & customer.iAvatarsize & """ width=""" & customer.iAvatarsize & """ border=""1"" style=""width:"&customer.iAvatarsize&";height:"&customer.iAvatarsize&";margin:0px;"
if not isLeeg(customer.sAvatarBorderColor) then
getClickAvatar=getClickAvatar& "border:1px solid " & customer.sAvatarBorderColor & """ "
else
getClickAvatar=getClickAvatar& "border-style:none"" "
end if
dim h
set h=new md5
getClickAvatar=getClickAvatar& "src=""http://www.gravatar.com/avatar.php?gravatar_id=" & h.hash(sEmail) & "&amp;default=" & customer.sQSUrl & "/fixedImages/avatar.jpg"" alt=""avatar"" alt=""title"" /></a>"
set h=nothing
end if
end function
public property get sEmailorNickname
select case convertGetal(customer.iLoginMode)
case 0
sEmailorNickname=sEmail
case 1
sEmailorNickname=sNickname
end select
end property
end class%>
