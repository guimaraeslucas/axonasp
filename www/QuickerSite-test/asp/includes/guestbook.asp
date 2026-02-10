
<%class cls_guestbook
Public iId
Public sCode
Public sName
Public dOnlineFrom
Public dOnlineUntil
Public sTemplate
Public sTemplateForm
Public sTemplateReply
Public bRequireValidation
Public sEmail
Public sFullTemplate
Public sWarningApproval
Public sBlockIP
Public sSortBy
Public sTemplateErr
Public iPaging, overruleCID,bDoSave
Private Sub Class_Initialize
iId=null
iPaging=10
sSortBy="recentfirst"
bRequireValidation=true
sEmail=customer.webmasterEmail
sFullTemplate="<div style=""width:95%"">" & vbcrlf & "<div>{ITEMS}</div>" & vbcrlf & "<div style=""width:100%;text-align:center"">{PAGING}</div>" & vbcrlf & "<div>{FORM}</div>" & vbcrlf & "</div>"
sTemplateForm="<hr /><h3>Leave a message:</h3>{TEMPLATEERROR}<table border=""0"" style=""border-style:none""><tr><td>Your name:*</td><td><input required style=""width:100%"" maxlength=""50"" type=""text"" value=""{AUTHOR}"" name=""author""/></td></tr><tr><td>Your email:</td><td><input style=""width:100%"" maxlength=""50"" type=""email"" value=""{EMAIL}"" name=""email""/></td></tr><tr><td>Your message:*</td><td><textarea required id=""{ID}"" name=""{ID}"" style=""width:100%"" cols=""40"" rows=""8"">{MESSAGE}</textarea></td></tr><tr><td> </td><td>{SMILEYS}</td></tr><tr><td>Are you human?</td><td>{CAPTCHA} <input style=""width:63px"" type=""text"" name=""captcha"" maxlength=""4"" size=""6"" /></td></tr><tr><td> </td><td>(*) " & l("mandatory") & "</td></tr><tr><td> </td><td><input type=""submit"" class=""art-button"" value=""Submit"" name=""post"" /> <input class=""art-button"" type=""reset"" value=""Reset"" name=""reset"" /></td></tr></table>"
sTemplate="<blockquote>"&vbcrlf&"<div><b>{AUTHOR}</b> " & l("says") & " ({TIMESTAMP}):</div>" & vbcrlf &  "<div>{MESSAGE}</div>" & vbcrlf & "{TEMPLATEREPLY}{WARNINGAPPROVAL}" & vbcrlf & "</blockquote>"
sTemplateReply="<div style=""margin:10px;padding:10px;background-color:#CDCDCD"">Reply from admin: {REPLY}</div>"
sWarningApproval="<div style=""color:Green""><strong>Thank you for posting! Your entry needs to be approved by the webmaster.</strong></div>"
sTemplateErr="<div style=""margin:10px;padding:10px;background-color:Red;color:White"">{ERROR}</div>"
pick(decrypt(request("iGBID")))
bDoSave=true
end sub
Public function getByCode(cValue)
on error resume next
getByCode=false
dim rs
set rs=db.execute("select iId from tblGuestbook where iCustomerID=" & cid & " and sCode='"& left(cleanup(cValue),50) & "'")
if not rs.eof then
pick(rs(0))
getByCode=true
end if
on error goto 0
end function
Public Function Pick(id)
dim sql, RS, sValue
if isNumeriek(id) then
sql = "select * from tblGuestbook where iCustomerID="&cid&" and iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sCode	= rs("sCode")
sName	= rs("sName")
dOnlineFrom	= rs("dOnlineFrom")
dOnlineUntil	= rs("dOnlineUntil")
sTemplate	= rs("sTemplate")
sTemplateForm	= rs("sTemplateForm")
bRequireValidation	= rs("bRequireValidation")
sEmail	= rs("sEmail")
sFullTemplate	= rs("sFullTemplate")
sWarningApproval	= rs("sWarningApproval")
sBlockIP	= rs("sBlockIP")
sTemplateReply	= rs("sTemplateReply")
sSortBy	= rs("sSortBy")
sTemplateErr	= rs("sTemplateErr")
iPaging	= rs("iPaging")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sCode) then
check=false
message.AddError("err_mandatory")
exit function
end if
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
exit function
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
rs.Open "select * from tblGuestBook where 1=2"
rs.AddNew
else
rs.Open "select * from tblGuestBook where iId="& iId
end if
rs("sCode")	= ucase(sCode)
rs("sName")	= sName
rs("dOnlineFrom")	= dOnlineFrom
rs("dOnlineUntil")	= dOnlineUntil
rs("sTemplate")	= sTemplate
rs("sTemplateForm")	= sTemplateForm
if isLeeg(overruleCID) then
rs("iCustomerID")	= cId
else
rs("iCustomerID")	= overruleCID
end if
rs("bRequireValidation")	= bRequireValidation
rs("sEmail")	= sEmail
rs("sFullTemplate")	= sFullTemplate
rs("sWarningApproval")	= sWarningApproval
rs("sBlockIP")	= sBlockIP
rs("sTemplateReply")	= sTemplateReply
rs("sSortBy")	= sSortBy
rs("sTemplateErr")	= sTemplateErr
rs("iPaging")	= iPaging
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function getRequestValues()
sCode	= replace(lcase(convertStr(Request.Form ("sCode")))," ","")
sName	= Request.Form ("sName")
dOnlineFrom	= convertDateFromPicker(Request.Form ("dOnlineFrom"))
dOnlineUntil	= convertDateFromPicker(Request.Form ("dOnlineUntil"))
sTemplate	= Request.Form ("sTemplate")
sTemplateForm	= Request.Form ("sTemplateForm")
bRequireValidation	= convertBool(Request.Form ("bRequireValidation"))
sEmail	= Request.Form ("sEmail")
sFullTemplate	= Request.Form ("sFullTemplate")
sWarningApproval	= Request.Form ("sWarningApproval")
sBlockIP	= Request.Form ("sBlockIP")
sTemplateReply	= Request.Form ("sTemplateReply")
sSortBy	= Request.Form ("sSortBy")
sTemplateErr	= Request.Form ("sTemplateErr")
iPaging	= convertGetal(Request.Form ("iPaging"))
end function
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("delete from tblGuestbookItem where iGuestbookID="& convertGetal(iId))
set rs=nothing
set rs=db.execute("delete from tblGuestbook where iId="& convertGetal(iId))
set rs=nothing
end if
end function
Public function copy()
if isNumeriek(iId) then
iId=null
sCode=GeneratePassword
save()
end if
end function
public function build()
if convertGetal(iId)=0 then
exit function
end if
dim iPage
iPage=convertGetal(left(Request.QueryString ("iPage"),4))
if iPage=0 then iPage=1
dim gform, showThis, eform
gform=sTemplateForm
eform=sTemplateErr
gform=replace(gform,"{SMILEYS}",shortListOfSmilies("gbform." & encrypt(iId)),1,-1,1)
gform=replace(gform,"{CAPTCHA}","<img style='vertical-align: middle;margin:0px 0px 0px 0px;border-style:none' alt='CAPTCHA' src=""" & C_DIRECTORY_QUICKERSITE & "/asp/includes/captcha.asp"" />",1,-1,1)
gform=replace(gform,"{ID}",encrypt(iId),1,-1,1)
gform="<a name=""form" & encrypt(iId) & """></a><form method=""post"" action=""default.asp?iId=" & encrypt(selectedpage.iId) & "#form"&encrypt(iId)&""" name=""gbform"" id=""gbform""><input type=""hidden"" value=""" & quotrep(request("pageAction")) & """ name=""pageAction"" /><input type=""hidden"" value=""" & quotrep(request("iItemID")) & """ name=""iItemID"" /><input type=""hidden"" value=""" & true & """ name=""postbackGB"" />" & gform & "</form>"
if convertBool(Request.Form ("postbackGB")) then
dim bgIsError
bgIsError=isleeg(session("CAPTCHA")) or LCase(session("CAPTCHA")) <> LCase(Left(Request.Form("CAPTCHA"),4))	or isLeeg(Request.Form("CAPTCHA"))
If bgIsError Then
eform=replace(eform,"{ERROR}",l("err_captcha"),1,-1,1)
gform=replace(gform,"{MESSAGE}",sanitize(Request.Form (encrypt(iId))),1,-1,1)
gform=replace(gform,"{AUTHOR}",sanitize(Request.Form ("AUTHOR")),1,-1,1)
gform=replace(gform,"{EMAIL}",sanitize(Request.Form ("email")),1,-1,1)
else
'save!
dim newEntry
set newEntry=new cls_guestbookitem
newEntry.sValue	= Request.Form (encrypt(iId))
newEntry.sMessageBy	= Request.Form ("author")
newEntry.sMessageByEmail	= Request.Form ("email")
newEntry.iGuestBookID	= iId
newEntry.ip	= UserIP()
if bDoSave then
if not newEntry.save() then
eform=replace(eform,"{ERROR}",l("err_mandatory"),1,-1,1)
gform=replace(gform,"{MESSAGE}",sanitize(Request.Form (encrypt(iId))),1,-1,1)
gform=replace(gform,"{AUTHOR}",sanitize(Request.Form ("AUTHOR")),1,-1,1)
gform=replace(gform,"{EMAIL}",sanitize(Request.Form ("email")),1,-1,1)
else
gform=replace(gform,"{MESSAGE}","",1,-1,1)
gform=replace(gform,"{AUTHOR}","",1,-1,1)
gform=replace(gform,"{EMAIL}","",1,-1,1)
session("GBI" & newEntry.iId)=true
showThis=convertGetal(newEntry.iId)
end if
end if
set newEntry=nothing
end if
end if
gform=replace(gform,"{MESSAGE}","",1,-1,1)
gform=replace(gform,"{AUTHOR}","",1,-1,1)
gform=replace(gform,"{EMAIL}","",1,-1,1)
if eform=sTemplateErr then
gform=replace(gform,"{TEMPLATEERROR}","",1,-1,1)
else
gform=replace(gform,"{TEMPLATEERROR}",eform,1,-1,1)
end if
dim gFull,gItemsRS,gItem,fullGitems,bApproved,gItemID,gItemReply,sql
sql="select * from tblGuestBookItem where iGuestBookID="& iId & " order by dCreatedTS "
if sSortby="recentfirst" then
sql=sql& " desc "
else
sql=sql& " asc "
end if
set gItemsRS=db.execute(sql)
dim recordCounter,doContinue
recordCounter=0
doContinue=true
while not gItemsRS.eof and doContinue
recordCounter=recordCounter+1
if recordCounter<=iPage*iPaging and recordCounter>(iPage*iPaging)-iPaging then
gItemID=gItemsRS("iId")
gItem=sTemplate
gItemReply=sTemplateReply
if not isLeeg(gItemsRS("sReply")) then
gItemReply=replace(gItemReply,"{REPLY}",linkUrls(sanitize(convertStr(gItemsRS("sReply")))),1,-1,1)
else
gItemReply=""
end if
gItem=replace(gItem,"{AUTHOR}",sanitize(gItemsRS("sMessageBy")),1,-1,1)
gItem=replace(gItem,"{MESSAGE}",addSmilies(linkUrls(sanitize(gItemsRS("sValue")))),1,-1,1)
gItem=replace(gItem,"{TIMESTAMP}",formatTimeStamp(gItemsRS("dCreatedTS")),1,-1,1)
gItem=replace(gItem,"{TEMPLATEREPLY}",gItemReply,1,-1,1)
bApproved=convertBool(gItemsRS("bApproved"))
if not bApproved then
gItem=replace(gItem,"{WARNINGAPPROVAL}",sWarningApproval,1,-1,1)
recordCounter=recordCounter-1
else
gItem=replace(gItem,"{WARNINGAPPROVAL}","",1,-1,1)
end if
gItem="<a name='gb"& gItemID &"'></a>" & gItem
if bApproved then
fullGitems=fullGitems & gItem
elseif showThis=convertGetal(gItemID) then
fullGitems=fullGitems & gItem
elseif convertBool(session("GBI" & convertGetal(gItemID))) then
fullGitems=fullGitems & gItem
end if
end if
if recordCounter>iPage*iPaging then
doContinue=false
end if
gItemsRS.movenext
wend
set gItemsRS=nothing
dim rsCount,pagingNav
set rsCount=db.execute("select count(*) from tblGuestBookItem where iGuestBookID="& iId)
rsCount=convertGetal(rsCount(0))
dim pb
for pb=1 to (rsCount/iPaging)+1
if pb=iPage then
pagingNav=pagingNav & pb & " "
else
pagingNav=pagingNav & "<a href=""default.asp?pageAction=" & quotrep(request("pageAction")) &"&amp;iItemID=" & quotrep(request("iItemID")) & "&amp;iId=" & encrypt(selectedPage.iId) & "&amp;iPage=" & pb & "#gb" & encrypt(iId) & """>" & pb & "</a> "
end if
next
if trim(pagingNav)="1" then pagingNav=""
gFull=sFullTemplate
if isBetween(dOnlineFrom,date(),dOnlineUntil) then
gFull=replace(gFull,"{FORM}",gform,1,-1,1)
else
gFull=replace(gFull,"{FORM}","",1,-1,1)
end if
gFull=replace(gFull,"{ITEMS}",fullGitems,1,-1,1)
gFull=replace(gFull,"{PAGING}",pagingNav,1,-1,1)
build="<a name='gb"&encrypt(iId)&"'></a>" & gFull
end function
public function entries
set entries=server.CreateObject ("scripting.dictionary")
dim rs, sql, itemID,itemAuthor,itemValue,itemApproved, itemIP, itemReply, itemEmail
sql="select * from tblGuestBookItem where iGuestBookID="& iId & " order by dCreatedTS desc"
set rs=db.execute(sql)
while not rs.eof
itemID	= rs("iId")
itemAuthor	= rs("sMessageBy")
itemValue	= rs("sValue")
itemApproved	= rs("bApproved")
itemIP	= rs("ip")
itemReply	= rs("sReply")
itemEmail	= rs("sMessageByEmail")
entries.Add itemID,itemAuthor & QS_VBScriptIdentifier & itemValue & QS_VBScriptIdentifier & itemApproved & QS_VBScriptIdentifier & itemIP & QS_VBScriptIdentifier & itemReply & QS_VBScriptIdentifier & itemEmail
rs.movenext
wend 
set rs=nothing
end function
public function copyToCustomer(oCID)
if isNumeriek(iId) then
dim oldId
oldID=iId
overruleCID	= oCID
iId=null
save()
end if
end function
end class%>
