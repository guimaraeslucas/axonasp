
<%dim message
set message = New cls_Messages
dim messageKeys,iM
if not isLeeg(Request.QueryString ("strMessage")) then
messageKeys=split(Request.QueryString ("strMessage"),",")
for iM=lbound(messageKeys) to ubound(messageKeys) 
if not isLeeg(messageKeys(iM)) then
message.addError(messageKeys(iM))
end if
next
end if
if not isLeeg(Request.QueryString ("fbMessage")) then
messageKeys=split(Request.QueryString ("fbMessage"),",")
for iM=lbound(messageKeys) to ubound(messageKeys) 
if not isLeeg(messageKeys(iM)) then
message.add(messageKeys(iM))
end if
next
end if
Class cls_Messages
private keys
private errors
private AllMessages
private bSticky
Sub Class_Initialize
Set keys = Server.CreateObject("scripting.dictionary")
Set errors = Server.CreateObject("scripting.dictionary")
set AllMessages = Server.CreateObject("scripting.dictionary")
bSticky=false
end sub
Sub class_terminate
Set keys = nothing
Set errors = nothing
Set AllMessages = nothing
end sub
Public function add(byval key)
keys(key) = "temporary value"
end function
Public function adderror(byval key)
errors(key) = "temporary value"
end function
Public function hasMessages()
hasMessages = (keys.Count > 0) or (errors.Count > 0)
end function
Public function hasErrors()
hasErrors = (errors.Count > 0)
end function
Public function showAlert()
loadMessages()
dim messages, message
set messages=server.CreateObject ("scripting.dictionary")
for each message in AllMessages
messages(message) = AllMessages(message)
next
dim key_number
For each key_number in keys
if allMessages.Exists (key_number) then
showAlert = showAlert & "<li>" & messages(key_number) & "</li>"
else
keys.Remove (key_number)
end if
select case key_number
case "fb_emailFound","fb_activationlink","fb_explvalidation","fb_templateadded"
bSticky=true
end select
Next
For each key_number in errors
if allMessages.Exists (key_number) then
showAlert = showAlert & "<li>" & messages(key_number) & "</li>"
else
errors.Remove (key_number)
end if
bSticky=true
Next
if errors.Count>0 or keys.count>0 then
showAlert=replace(showAlert,"\n","<br />",1,-1,1)
showAlert=replace(showAlert,"\'","'",1,-1,1)
set cPopup=new cls_popup
cPopup.bUseAsSystemMessage=true
if hasErrors then
cPopup.messageMode="err"
showAlert="<div style=""display:none""><div id=""errMessage"" style=""height:130px;background-image: url('" & C_DIRECTORY_QUICKERSITE & "/fixedImages/error_bg.jpg');background-repeat: no-repeat;background-position:right bottom;font-size:10pt;font-family:Arial;color:#666;font-weight:700;padding:10px;background-color:#fff;""><ul>" &   showAlert & "</ul></div></div>"
else
cPopup.messageMode="fb"
showAlert="<div style=""display:none""><div id=""fbMessage"" style=""height:130px;background-image: url('" & C_DIRECTORY_QUICKERSITE & "/fixedImages/success_bg.jpg');background-repeat: no-repeat;background-position:right bottom;font-size:10pt;font-family:Arial;color:#666;font-weight:700;padding:10px;background-color:#fff;""><ul>" &  showAlert & "</ul></div></div>"
end if
cPopup.bSticky=bSticky
end if
end function
Private function loadMessages()
AllMessages.RemoveAll 
AllMessages.Add "err_mandatory",l("err_mandatory")
AllMessages.Add "err_email",l("err_email")
AllMessages.Add "err_pw",l("err_pw")
AllMessages.Add "err_login",l("err_login")
AllMessages.Add "err_makeChoice",l("err_makeChoice")
AllMessages.Add "err_emailNotFound",l("err_emailNotFound")
AllMessages.Add "err_ICO_file",l("err_ICO_file")
AllMessages.Add "err_ICO_fileSize",l("err_ICO_fileSize")
AllMessages.Add "err_newFile",l("err_newFile")
AllMessages.Add "err_JPGGIF_file",l("err_JPGGIF_file")
AllMessages.Add "err_LOGOFileSize",l("err_LOGOFileSize")
AllMessages.Add "err_fileSize",l("err_fileSize")
AllMessages.Add "err_fileType",l("err_fileType")
AllMessages.Add "err_appCode",l("err_appCode")
AllMessages.Add "err_labelCode",l("err_labelCode")
AllMessages.Add "err_url",l("err_url")
AllMessages.Add "err_reg_or_fb",l("err_reg_or_fb")
AllMessages.Add "err_catorform",l("err_catorform")
AllMessages.Add "err_captcha",l("err_captcha")
AllMessages.Add "err_feed",l("err_feed")
AllMessages.Add "err_constant",l("err_constant")
AllMessages.Add "err_newfolder",l("err_newfolder")
AllMessages.Add "err_fileexists",l("err_fileexists")
AllMessages.Add "err_fileupload",l("err_fileupload")
AllMessages.Add "err_backsitepw",l("err_backsitepw")
AllMessages.Add "err_doublefeed",l("err_doublefeed")
AllMessages.Add "err_ufl",l("err_ufl")
AllMessages.Add "pwnomatch",l("pwnomatch")
AllMessages.Add "err_ufl_an",l("err_ufl_an")
AllMessages.Add "err_foldernotexists",l("err_foldernotexists")
AllMessages.Add "err_foldernotallowed",l("err_foldernotallowed")
AllMessages.Add "err_activationlink",l("err_activationlink")
AllMessages.Add "err_doubleemail",l("err_doubleemail")
AllMessages.Add "err_wrongactivationlink",l("err_wrongactivationlink")
AllMessages.Add "err_mandatoryprofile",l("err_mandatoryprofile")
AllMessages.Add "err_doublenickname",l("err_doublenickname")
AllMessages.Add "err_firstandsecondpassword",l("err_firstandsecondpassword")
AllMessages.Add "err_topictoolong",l("err_topictoolong")
AllMessages.Add "err_listpagepurl", l("err_listpagepurl")
AllMessages.Add "err_numeric", "Only numbers for amounts!"
AllMessages.Add "fb_saveOK",l("fb_saveOK")
AllMessages.Add "fb_emailFound",l("fb_emailFound")
AllMessages.Add "fb_fileremoved",l("fb_fileremoved")
AllMessages.Add "fb_folderremoved",l("fb_folderremoved")
AllMessages.Add "fb_folderadded",l("fb_folderadded")
AllMessages.Add "fb_fileadded",l("fb_fileadded")
AllMessages.Add "fb_activationlink",l("fb_activationlink")
AllMessages.Add "fb_passwordreset",l("fb_passwordreset")
AllMessages.Add "fb_topicremoved",l("fb_topicremoved")
AllMessages.Add "fb_replyremoved",l("fb_replyremoved")
AllMessages.Add "fb_replysave",l("fb_replysave")
AllMessages.Add "fb_topicsave",l("fb_topicsave")
AllMessages.Add "fb_subscribetheme",l("fb_subscribetheme")
AllMessages.Add "fb_subscribetopic",l("fb_subscribetopic")
AllMessages.Add "fb_subscriptionthemecancelled",l("fb_subscriptionthemecancelled")
AllMessages.Add "fb_subscriptiontopiccancelled",l("fb_subscriptiontopiccancelled")
AllMessages.Add "fb_activationlinkresend",l("fb_activationlinkresend")
AllMessages.Add "fb_replyvalidated",l("fb_replyvalidated")
AllMessages.Add "fb_topicvalidated",l("fb_topicvalidated")
AllMessages.Add "fb_explvalidation",l("fb_explvalidation")
AllMessages.Add "fb_templateadded",l("fb_templateadded")
end function
end Class%>
