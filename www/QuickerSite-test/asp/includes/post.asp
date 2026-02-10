
<%class cls_post
Public iId, dCreatedTS, dUpdatedTS, iContactID, iThemeID, iPostID, sSubject, sBody, sKey, bNeedsToBeValidated, sAnName, sFilename, sFileDesc, iFileSize, bRemoveAnyway
private p_theme
private p_parentTopic
Private Sub Class_Initialize
On Error Resume Next
set p_theme=nothing
set p_parentTopic=nothing
bRemoveAnyway=false
bNeedsToBeValidated=false
pick(decrypt(request("iPostID")))
On Error Goto 0
end sub
Public Function Pick(id)
if isNumeriek(id) then
dim sql, RS, postArr(14)
if not isArray(Application("POST"&id))  then
sql = "SELECT tblPost.* FROM tblTheme INNER JOIN tblPost ON tblTheme.iId = tblPost.iThemeID WHERE tblTheme.iCustomerID=" & cId & " AND tblPost.iId="& id
set RS = db.execute(sql)
if not rs.eof then
postArr(0)	= rs("iId")
postArr(1)	= rs("dCreatedTS")
postArr(2)	= rs("dUpdatedTS")
postArr(3)	= rs("iContactID")
postArr(4)	= rs("iThemeID")
postArr(5)	= rs("iPostID")
postArr(6)	= rs("sSubject")
postArr(7)	= rs("sBody")
postArr(8)	= rs("sKey")
postArr(9)	= rs("bNeedsToBeValidated")
postArr(10) = rs("sAnName")
postArr(11)	= rs("sFilename")
postArr(12)	= rs("sFileDesc")
postArr(13)	= rs("iFileSize")
Application("POST"&id)=postArr
end if
set RS = nothing
end if
iId	= Application("POST"&id)(0)
dCreatedTS	= Application("POST"&id)(1)
dUpdatedTS	= Application("POST"&id)(2)
iContactID	= Application("POST"&id)(3)
iThemeID	= Application("POST"&id)(4)
iPostID	= Application("POST"&id)(5)
sSubject	= Application("POST"&id)(6)
sBody	= Application("POST"&id)(7)
sKey	= Application("POST"&id)(8)
bNeedsToBeValidated	= Application("POST"&id)(9)
sAnName	= Application("POST"&id)(10)
sFilename	= Application("POST"&id)(11)
sFileDesc	= Application("POST"&id)(12)
iFileSize	= Application("POST"&id)(13)
end if
end function
private function bCanModify
on error resume next
bCanModify=true
if convertGetal(logon.contact.iId)=0 then
bCanModify=false
exit function
end if
if convertGetal(theme.iContactID)=convertGetal(logon.contact.iId) then
exit function
end if
if convertGetal(logon.contact.iId)<>convertGetal(iContactID) then
bCanModify=false
exit function
end if
if logon.contact.istatus<cs_write then
bCanModify=false
exit function
end if
if cdate(dCreatedTS)<dateAdd("h",-12,now()) then
bCanModify=false
exit function
end if
if err.number <> 0 then bCanModify=false
on error goto 0
end function
public function removeByKey(key)
dim rs
set rs=db.execute("select iId from tblPost where sKey='"& left(cleanup(key),128) &"'")
if not rs.eof then
pick(rs(0))
remove()
end if
set rs=nothing
end function
Public Function Check()
Check = true
if isLeeg(sSubject) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sBody) and isLeeg(sFilename) then
if not bRemoveAnyway then
check=false
message.AddError("err_mandatory")
end if
end if
if len(sBody)>200000 then
check=false
message.AddError("err_topictoolong")
end if
End Function
Public Function Save
if check() then
save=true
else
save=false
exit function
end if
dim rs, isNew
isNew=false
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblPost where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
isNew=true
if convertBool(theme.bValidation) then
if convertGetal(theme.iContactID)<>convertGetal(logon.contact.iId) then
bNeedsToBeValidated=true
end if
end if
else
rs.Open "select * from tblPost where iId="& iId
end if
rs("dUpdatedTS")	= now()
if (iContactID<>logon.contact.iId and convertGetal(iContactID)<>0) or not isLeeg(sAnName) then
rs("iContactID")	= iContactId 'oorsponkelijke autheyr bewaren
else
rs("iContactID")	= logon.contact.iId
end if
rs("iThemeID")	= iThemeID
rs("iPostID")	= iPostID
rs("sSubject")	= sSubject
rs("sBody")	= sBody
rs("bNeedsToBeValidated")	= bNeedsToBeValidated
rs("sAnName")	= sAnName
rs("sFilename")	= sFilename
rs("sFileDesc")	= sFileDesc
rs("iFileSize")	= convertGetal(iFileSize)
'contactID initialiseren
iContactID	= rs("iContactID")
dUpdatedTS=now()
rs.Update 
iId = convertGetal(rs("iId"))
Application("POST"&iId)=""
rs.close
Set rs = nothing
if isNew then
sKey=sha256(iId&GeneratePassWord&GeneratePassWord)
Save
end if
'clear cache
dim themePage
set rs=db.execute("select iId from tblPage where iCustomerID=" & cId & " and sValue like '%"&theme.sCode&"]%' OR iThemeID="& iThemeID)
while not rs.eof
set themePage=new cls_page
themePage.pick(rs(0))
themePage.clearRSScache()
set themePage=nothing
rs.movenext
wend
set rs=nothing
'send emails
dim theMail, mailSQL
if isNew then 'zowiezo alleen mails sturen als er nieuwe berichten zijn, niet bij updates...
if isLeeg(sBody) then
'sBody="Download Attachment"
end if
'mails versturen naar thema-inschrijvingen
if theme.iSubLevel=QS_theme_sublevel_theme then
mailSQL="select tblContact.sEmail from tblContact "
mailSQL=mailSQL& " INNER JOIN tblThemeSubscription on  tblThemeSubscription.iContactID=tblContact.iId "
mailSQL=mailSQL& " WHERE tblThemeSubscription.iThemeID="& convertGetal(theme.iId)
mailSQL=mailSQL &" and iContactID<>"& convertGetal(logon.contact.iId)
mailSQL=mailSQL& " and tblContact.sEmail<>'' and tblContact.sEmail is not null"
set rs=db.execute(mailSQL)
while not rs.eof
set theMail=new cls_mail_message
if convertGetal(iPostID)<>0 then 'reply!
theMail.subject=theme.sSubjectNotification
theMail.body=preparenotifyReply(sBody,sFileLink(sFileName,sFileDesc,iFileSize))
else
theMail.subject=theme.sTopicSubjectNotification
theMail.body=preparenotifyTopic(sBody,sFileLink(sFileName,sFileDesc,iFileSize))
end if
theMail.receiver=convertStr(rs(0))
theMail.send
rs.movenext
set theMail=nothing
wend
set rs=nothing
end if
'mails versturen naar topic-inschrijvingen
if theme.iSubLevel>QS_theme_sublevel_none then
if convertGetal(iPostID)<>0 then 'het gaat om een reply
mailSQL="select tblContact.sEmail from tblContact "
mailSQL=mailSQL& " INNER JOIN tblThemeTopicSubscription on  tblThemeTopicSubscription.iContactID=tblContact.iId "
mailSQL=mailSQL& " WHERE tblThemeTopicSubscription.iPostID="& iPostID
mailSQL=mailSQL &" and iContactID<>"& convertGetal(logon.contact.iId)
mailSQL=mailSQL& " and tblContact.sEmail<>'' and tblContact.sEmail is not null"
set rs=db.execute(mailSQL)
while not rs.eof
set theMail=new cls_mail_message
theMail.subject=theme.sSubjectNotification
theMail.body=preparenotifyReply(sBody,sFileLink(sFileName,sFileDesc,iFileSize))
theMail.receiver=convertStr(rs(0))
theMail.send
rs.movenext
set theMail=nothing
wend
set rs=nothing
end if
end if
'send notification to moderator? alleen als hij de post zelf niet deed...
if theme.bForwardPostsToModerator and logon.contact.iId<>theme.iContactID then
set theMail=new cls_mail_message
theMail.receiver=theme.contact.sEmail
if convertGetal(iPostID)<>0 then 'reply!
theMail.subject=l("moderatornotification") & " - " & theme.sSubjectNotification
theMail.body=preparenotifyReply(sBody,sFileLink(sFileName,sFileDesc,iFileSize)) & "<hr />" & getVisitorDetails
else
theMail.subject=l("moderatornotification") & " - " & theme.sTopicSubjectNotification
'response.end
theMail.body=preparenotifyTopic(sBody,sFileLink(sFileName,sFileDesc,iFileSize))& "<hr />" & getVisitorDetails
end if
theMail.send
set theMail=nothing
end if
else
'niet nieuw, dus een update!, mogelijks een update van forum
dim RSU
set RSU=db.execute("update tblPost set iThemeID=" & convertGetal(iThemeID) & " where iPostID=" & convertGetal(iId))
set RSU=nothing
end if
selectedPage.clearPageCache()
end function
public function getRequestValues()
sSubject	= convertStr(Request.Form("sSubject"))
sBody	= removeEmptyP(convertStr(Request.Form("sBody")))
sAnName	= convertStr(Request.Form("sAnName"))
end function
public function remove
if not isLeeg(iId) then
removeATT()
 
dim rs
set rs=db.execute("delete from tblThemeTopicSubscription where iPostId="& iId)
set rs=nothing
set rs=db.execute("delete from tblPost where iPostID="& iId)
set rs=nothing
set rs=db.execute("delete from tblPost where iId="& iId)
set rs=nothing
theme.clearRSScache()
Application("POST"&iId)=""
end if
end function
public function contact
set contact=getContact(iContactID)
contact.sAnName=sAnname
end function
public function theme
if p_theme is nothing then
set p_theme=new cls_theme
p_theme.pick(iThemeID)
end if
set theme=p_theme
end function
public function buildShortPost
'add security here
'bNeedsToBeValidated
if theme.iSearchType=0 then
redirectToHP()
elseif theme.iSearchType=2 and not logon.authenticatedIntranet then
redirectToHP()
elseif bNeedsToBeValidated then
redirectToHP()
end if
buildShortPost="<div id='QS_theme' style='margin-top:20px;"
if selectedPage.iThemeID=theme.iId then
buildShortPost=buildShortPost &"width:"&themeObj.iWidth&";"
else
'buildShortPost=buildShortPost &"word-break: break-all;"
end if
if not isLeeg(theme.sColorEven) or not isLeeg(theme.sColorUnEven) then
buildShortPost=buildShortPost & "background-color:" & theme.sColorEven
end if
buildShortPost=buildShortPost & "'"
buildShortPost=buildShortPost&">"
buildShortPost=buildShortPost&"<div class='QS_theme_subject'>"
buildShortPost=buildShortPost&"<span style='width:6px'> </span>" & linkUrls(quotrep(sSubject))
buildShortPost=buildShortPost&"</div>" 'end title
buildShortPost=buildShortPost&"<div id='topicTable"&encrypt(iId)&"' style=""display:block"">"
if theme.bAllowHTML then
buildShortPost=buildShortPost&"<div class='QS_theme_body'>"
buildShortPost=buildShortPost& filterJS(sBody) & "</div>"
else
buildShortPost=buildShortPost&"<div class='QS_theme_body'>"
if customer.bUseAvatars then
buildShortPost=buildShortPost& "<div style=""float:left;margin-top:3px;margin-right:13px;margin-bottom:11px;width:" & customer.iAvatarSize & "px"">" & contact.getAvatar() &"</div>"
end if
buildShortPost=buildShortPost&addSmilies(linkUrls(quotrep(sBody))) & "</div>"
end if
buildShortPost=buildShortPost&sFileLink(sFileName,sFileDesc,iFileSize)
buildShortPost=buildShortPost&"<div class='QS_theme_footer'>"&formatTimeStamp(dUpdatedTS)&" | " & contact.sClickNickName
if theme.bAllowComments  then 
'comments opvissen vooraleer je de footer sluit
dim cReplies,replyKey,styleReplies
set cReplies=replies(convertGetal(logon.contact.iId)=convertGetal(theme.iContactID))
if cReplies.Count>0 then
buildShortPost=buildShortPost&" | " & l("replies") & " ("&cReplies.count&")"
else
buildShortPost=buildShortPost&" | " & l("replies") & " (0)"
end if
styleReplies="block"
dim replyCounter,buildReply
replyCounter=1
if cReplies.Count>0 then
buildShortPost=buildShortPost&"<div id='replies"&encrypt(iId)&"' style='display:"&styleReplies&";text-align:left;font-size:0.9em'><a name=""comments""></a><ul>"
for each replyKey in cReplies
buildReply=""
buildReply=buildReply& "<li style='margin-top:15px'>"
buildReply=buildReply&"<a name='"&encrypt(replyKey)&"'></a><strong>" & cReplies(replyKey).contact.sClickNickName & " "& l("says") & ":</strong>&nbsp;"
buildReply=buildReply& "<span style='font-size:0.9em'>(" & formatTimeStamp(cReplies(replyKey).dUpdatedTS) & ")</span><div style='margin-top:5px'>"
if customer.bUseAvatars then
buildReply=buildReply& "<div style=""float:left;margin-top:3px;margin-right:13px;margin-bottom:11px;width:" & customer.iAvatarSize & "px"">" & cReplies(replyKey).contact.getAvatar() &"</div>"
end if
buildReply=buildReply& "" & addSmilies(linkUrls(sanitize(cReplies(replyKey).sBody)))
buildReply=buildReply & sFileLink(cReplies(replyKey).sFilename,cReplies(replyKey).sFileDesc,cReplies(replyKey).iFileSize)
buildReply=buildReply&"</div><div style=""clear:both""></div>"
buildReply=buildReply&"</li>"
buildShortPost=buildShortPost&buildReply
replyCounter=replyCounter+1
next
buildShortPost=buildShortPost&"</ul></div>"
end if
set cReplies=nothing
end if
buildShortPost=buildShortPost&"</div>"
'tonen of niet?
'if showCompactMode then
buildShortPost=buildShortPost&"</div>"
'end if
buildShortPost=buildShortPost&"</div>"
if theme.postCount=0 then
theme.postCount=1
else
theme.postCount=0
end if
buildShortPost=replace(buildShortPost,"[","|R|R|R|",1,-1,1)
end function 
public function buildPost(themeObj,showReplies)
if themeObj.iType=QS_theme_ts then
if convertGetal(logon.contact.iId)<>convertGetal(themeObj.iContactID) and convertGetal(logon.contact.iId)<>convertGetal(iContactID) and not convertGetal(iContactID)=0 then
exit function
end if
end if
buildPost="<div id='QS_theme' style='margin-top:20px;"
if selectedPage.iThemeID=themeObj.iId then
buildPost=buildPost &"width:"&themeObj.iWidth&";"
else
'buildPost=buildPost &"word-break: break-all;"
end if
if not isLeeg(themeObj.sColorEven) or not isLeeg(themeObj.sColorUnEven) then
buildPost=buildPost & "background-color:"
'decide bgcolor
if themeObj.postCount=0 then
buildPost=buildPost & themeObj.sColorEven
else
buildPost=buildPost & themeObj.sColorUnEven
end if
end if
buildPost=buildPost & "'"
buildPost=buildPost&"><a name="& """" & "apost" & encrypt(iId) & """" & "></a>"
buildPost=buildPost&"<div class='QS_theme_subject'>"
dim showCompactMode
showCompactMode=themeObj.bCompactList and not showReplies and not printReplies
if not isLeeg(iId) then showCompactMode=false
'tonen of niet?
if showCompactMode then
buildPost=buildPost&"<a href='#' "
buildPost=buildPost&"onclick=""javascript:return compact('"&encrypt(iId)&"')"">"
buildPost=buildPost&"<img alt='' style='margin:0px;border-style:none' src='" & C_DIRECTORY_QUICKERSITE & "/fixedImages/plus2.gif' id='zoom"&encrypt(iId)&"' /></a>"
else
buildPost=buildPost&"<a href='#' "
buildPost=buildPost&"onclick=""javascript:return uncompact('"&encrypt(iId)&"')"">"
buildPost=buildPost&"<img alt='' style='margin:0px;border-style:none' src='" & C_DIRECTORY_QUICKERSITE & "/fixedImages/minus2.gif' id='zoom"&encrypt(iId)&"' /></a>"
end if
buildPost=buildPost&"<span style='width:6px'> </span><a href='" & prepareLink & "'>"
if bNeedsToBeValidated then
buildPost=buildPost & l("tobevalidated")  & "&nbsp;-&nbsp;" & linkUrls(quotrep(sSubject)) & "</a>"
else
buildPost=buildPost & linkUrls(quotrep(sSubject)) & "</a>"
end if 
buildPost=buildPost&"</div>" 'end title
if showCompactMode then
buildPost=buildPost&"<div id='topicTable"&encrypt(iId)&"' style=""display:none"">"
else
buildPost=buildPost&"<div id='topicTable"&encrypt(iId)&"' style=""display:block"">"
end if
if themeObj.bAllowHTML then
buildPost=buildPost&"<div class='QS_theme_body'>" & filterJS(sBody) & "</div>"
else
buildPost=buildPost&"<div class='QS_theme_body'>"
if customer.bUseAvatars then
buildPost=buildPost& "<div style=""float:left;margin-top:3px;margin-right:13px;margin-bottom:11px;width:" & customer.iAvatarSize & "px"">" & contact.getAvatar() &"</div>"
end if
buildPost=buildPost& addSmilies(linkUrls(quotrep(sBody))) & "</div>"
end if
buildPost=buildPost&sFileLink(sFileName,sFileDesc,iFileSize)
'if convertGetal(themeObj.iContactID)=convertGetal(logon.contact.iId) then 'allow removal
'	buildPost=buildPost="<div><a href="""">Delete attachment</a></div>"
'end if
buildPost=buildPost&"<div class='QS_theme_footer'>"&formatTimeStamp(dUpdatedTS)&" | " & contact.sClickNickName
if not themeObj.bLocked then
if not logon.authenticatedIntranet then
buildPost=buildPost&" | <a href='default.asp?item="& request("item") &"&amp;iId="&encrypt(selectedPage.iId)&"&amp;iPostID="&encrypt(iId)&"&amp;pageAction="&cloginIntranet&"'>" & l("login") & "</a>"
else
if convertGetal(logon.contact.iStatus)>cs_silent then
buildPost=buildPost&" | "&l("loggedinas")&"&nbsp;<a href='default.asp?pageAction="& cProfile &"'>" & quotrep(logon.contact.sNickname) & "</a>"
else
buildPost=buildPost&" | "&l("loggedinas")&"&nbsp;" & quotrep(logon.contact.sNickname)
end if
if not isLeeg(customer.intranetLogOff) then
buildPost=buildPost&" | <a onclick=" & """" & "javascript:return confirm('"&l("areyousure")&"');" & """" & " href='default.asp?iId=" & encrypt(selectedPage.IId) & "&amp;pageAction="& cLogOff &"'>" & customer.intranetLogOff & "</a>"
end if
end if
if bCanModify then
buildPost=buildPost&" | <a href='default.asp?"&QS_secCodeURL&"&amp;iId="& encrypt(selectedPage.iId) &"&amp;iPostID="&encrypt(iId)&"&amp;themeAction="&cPostTopic&"&amp;item=" & Request("item") & "'>" & l("modify") & "</a>"
end if
end if
'comment
if (themeObj.allowSub and (logon.contact.istatus=>cs_write or convertBool(themeObj.bAllowAP))) or convertBool(themeObj.bAllowAP) then
if themeObj.iSubLevel>=QS_theme_sublevel_authortopic and not logon.contact.bHasSubscribedToTheme(themeObj.iId) and convertGetal(logon.contact.iId)<>0 then
if (themeObj.iSubLevel=QS_theme_sublevel_authortopic and iContactID=logon.contact.iId) or themeObj.iSubLevel>QS_theme_sublevel_authortopic then
if logon.contact.bHasSubscribedToTopic(iId) then
buildPost=buildPost&" | <a onclick=" & """" & "javascript:return confirm('"&l("areyousure")&"');" & """" & " href='default.asp?"&QS_secCodeURL&"&amp;iId="&encrypt(selectedpage.iId)&"&amp;themeAction="&cUnSubscribeToTopic&"&amp;iPostID="&encrypt(iId)&"&amp;item=" & Request("item") & "'>"&l("unsubscribetopic")&"</a>"
else
buildPost=buildPost&" | <a onclick=" & """" & "javascript:return confirm('"&l("expl_subscribe_topic")&"');" & """" & " href='default.asp?"&QS_secCodeURL&"&amp;iId="&encrypt(selectedpage.iId)&"&amp;themeAction="&cSubscribeToTopic&"&amp;iPostID="&encrypt(iId)&"&amp;item=" & Request("item") & "'>"&l("subscribetopic")&"</a>"
end if
end if
end if
if themeObj.bAllowComments then
 
if not isLeeg(Request.Form ("replyTopic"&encrypt(iId))) or not isLeeg(request.form("sFilename")) then
checkCSRF()
dim reply
set reply=new cls_post
reply.sFilename=""
reply.sFileDesc=""
reply.iFileSize=""
reply.iId=null 'initialiseren op null
reply.iPostID=iId
reply.iContactID=logon.contact.iId
reply.sSubject="COMMENT"
reply.iThemeID=iThemeID
if themeObj.bAllowAP then
reply.sAnName=left(Request.Form ("replySender"&encrypt(iId)),50)
else
reply.sAnName=""
end if
if not isLeeg(request.form("sFilename")) then
reply.sFilename=request.form("sFilename")
reply.sFileDesc=request.form("sFileDesc")
reply.iFileSize=request.form("iFileSize")
end if
reply.sBody=left(Request.Form ("replyTopic"&encrypt(iId)),(themeObj.iLimitReplyTo+50))
if themeObj.bAllowAP and not (logon.authenticatedIntranet and logon.contact.istatus=>cs_write) then
'additional checking for anonymous replies
if isLeeg(reply.sAnName) then
message.AddError("err_mandatory")
end if
if isLeeg(reply.sBody) then
message.AddError("err_mandatory")
end if

dim bgAIsError
bgAIsError=isleeg(session("CAPTCHA")) or LCase(session("CAPTCHA")) <> LCase(Left(Request.Form("CAPTCHA"),4))	or isLeeg(Request.Form("CAPTCHA"))

If bgAIsError Then
' Add mandatory error if security image was not correct
message.AddError("err_captcha")
end if
end if
if not message.hasErrors then
if reply.Save then
dim feedbackMessages
feedbackMessages="fb_replysave"
if reply.bNeedsToBeValidated then
feedbackMessages=feedbackMessages& ",fb_explvalidation"
end if
set reply=nothing
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId) &"&iPostID="&encrypt(iId)&"&fbMessage="&feedbackMessages&"&item=" & request("item"))
end if
end if 
set reply=nothing
end if
if Request.QueryString ("btnaction")=cRemoveReply then
checkCSRF()
dim removePost
set removePost=new cls_post
if convertGetal(themeObj.iContactID)=convertGetal(logon.contact.iId) then
'remove post
removePost.remove()
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId) &"&iPostID="& encrypt(removePost.iPostID)&"&fbMessage=fb_replyremoved&item=" & request("item"))
else
'in het riet sturen
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId))
end if
end if
if Request.QueryString ("btnaction")="removereplyatt" then
checkCSRF()
dim removeAttPost
set removeAttPost=new cls_post
if convertGetal(themeObj.iContactID)=convertGetal(logon.contact.iId) then
'remove post
removeAttPost.removeAtt()
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId) &"&iPostID="& encrypt(removeAttPost.iPostID)&"&fbMessage=fb_saveOK&item=" & request("item"))
else
'in het riet sturen
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId))
end if
end if
if Request.QueryString ("btnaction")=cValidateReply then
checkCSRF()
dim validatePost
set validatePost=new cls_post
if convertGetal(themeObj.iContactID)=convertGetal(logon.contact.iId) then
'validate post
validatePost.validate()
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId) &"&iPostID="& encrypt(validatePost.iPostID)&"&fbMessage=fb_replyvalidated&item=" & request("item"))
else
'in het riet sturen
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId))
end if
end if
if themeObj.bAllowAP and not (logon.authenticatedIntranet and logon.contact.istatus=>cs_write) then
buildPost=buildPost&" | <a href='default.asp?btreatasAN=1&amp;themeAction=" & cPostReply &"&amp;iId=" & encrypt(selectedPage.iId) & "&amp;iPostID=" & encrypt(iId) & "'>" & l("reply") & "</a>"
else
buildPost=buildPost&" | <a href='#' onclick="& """" & "javascript:if(document.getElementById('replyTable"&encrypt(iId)&"').style.display=='block'){document.getElementById('replyTable"&encrypt(iId)&"').style.display='none';return false;}else{document.getElementById('replyTable"&encrypt(iId)&"').style.display='block';replyForm" & encrypt(iId) & ".replyTopic"&encrypt(iId) &".focus();};return false;" & """" &">" & l("reply") & "</a>"
end if
end if
end if
if themeObj.bAllowComments  then 
'comments opvissen vooraleer je de footer sluit
dim cReplies,replyKey,styleReplies
set cReplies=replies(convertGetal(logon.contact.iId)=convertGetal(themeObj.iContactID))
if cReplies.Count>0 then
buildPost=buildPost&" | <a href='#' onclick=" & """" & "javascript:if(document.getElementById('replies"&encrypt(iId)&"').style.display=='block'){document.getElementById('replies"&encrypt(iId)&"').style.display='none';return false;}else{document.getElementById('replies"&encrypt(iId)&"').style.display='block'};return false;" & """" & ">" & l("replies") & " ("&cReplies.count&")</a>"
else
buildPost=buildPost&" | " & l("replies") & " (0)"
end if
if showReplies or printReplies then 
styleReplies="block"
else
styleReplies="none"
end if
if (logon.authenticatedIntranet and logon.contact.istatus=>cs_write and  not themeObj.bLocked and themeObj.bAllowComments) or (themeObj.bAllowComments and themeObj.bAllowAP) then
dim btreatasAN
btreatasAN=false
if themeObj.bAllowAP and not (logon.authenticatedIntranet and logon.contact.istatus=>cs_write) then
btreatasAN=true
end if
buildPost=buildPost&"<div style='display:"
if btreatasAN and convertBool(request("btreatasAN")) then
buildPost=buildPost&"block"
else
buildPost=buildPost&"none"
end if
buildPost=buildPost&"' id='replyTable"&encrypt(iId)&"'>"
buildPost=buildPost&"<form action='default.asp' method='post' name='replyForm"&encrypt(iId)&"'>"
buildPost=buildPost&QS_secCodeHidden
buildPost=buildPost&"<table border=""0"" style=""border-style:none"" align=""center"">"
if btreatasAN  then
buildPost=buildPost&"<tr>"
buildPost=buildPost&"<td style=""border-style:none"" align=""left"">"
buildPost=buildPost& themeObj.sLabelYourName & "<br/><input maxlength=""40"" type=""text"" id='replySender"&encrypt(iId)&"' name='replySender"&encrypt(iId)&"' size=""30""  value=""" & sanitize(request.form("replySender"&encrypt(iId))) & """ />"
buildPost=buildPost&"</td>"
buildPost=buildPost&"</tr>"
end if
buildPost=buildPost&"<tr>"
buildPost=buildPost&"<td style=""border-style:none"">"
buildPost=buildPost&"<input type='hidden' name='themeAction' value='"&cPostReply&"' />"
buildPost=buildPost&"<input type='hidden' name='iId' value='"&encrypt(selectedpage.iId)&"' />"
buildPost=buildPost&"<input type='hidden' name='item' value="""&quotrep(request("item"))&""" />"
buildPost=buildPost&"<input type='hidden' name='iPostID' value='"&encrypt(iId)&"' />"
buildPost=buildPost&"<input type='hidden' name='sFileName' value='' />"
buildPost=buildPost&"<input type='hidden' name='sFileDesc' value='' />"
buildPost=buildPost&"<input type='hidden' name='iFileSize' value='' />"
if convertBool(request("btreatasAN")) then
buildPost=buildPost&"<input type='hidden' name='btreatasAN' value='1' />"
end if
buildPost=buildPost&"<textarea onKeyDown=""javascript:textCounter(this.form.replyTopic"&encrypt(iId)&",'remLen" & encrypt(iId) & "',"&themeObj.iLimitReplyTo&");"" onKeyUp=""textCounter(javascript:this.form.replyTopic"&encrypt(iId)&",'remLen" & encrypt(iId) & "',"&themeObj.iLimitReplyTo&");"" id='replyTopic"&encrypt(iId)&"' name='replyTopic"&encrypt(iId)&"' style='width:100%' rows='10'>" & sanitize(request.form("replyTopic"&encrypt(iId))) & "</textarea>"
buildPost=buildPost&"</td>"
buildPost=buildPost&"</tr>"
buildPost=buildPost&"<tr><td style=""border-style:none""><span id='remLen" & encrypt(iId) & "'>" & themeObj.iLimitReplyTo & "</span>&nbsp;"& l("charactersleft") 
if convertBool(themeObj.bFileUploads) and convertBool(logon.authenticatedIntranet) then
buildPost=buildPost&" - <span id=""uSS" & encrypt(iId) & """><a href=""default.asp?pageAction=fileupload&amp;RP=1&amp;iPostID=" & encrypt(iId) & """ class=""QSPPAVATAR"">Attach file</a></span>"
end if
buildPost=buildPost& "</td></tr>"
buildPost=buildPost&"<tr>"
buildPost=buildPost&"<td style=""border-style:none"">" 
if convertBool(themeObj.bSmileys) then
buildPost=buildPost& "<div style='margin:2px'>" & shortListOfSmilies("replyForm" & encrypt(iId) & ".replyTopic" & encrypt(iId)) & "</div>"
end if
if btreatasAN  then
buildPost=buildPost&"</td></tr><tr><td style=""border-style:none"" align=""left""><table style=""border-style:none""><tr><td style=""border-style:none"">" & l("captcha") & "</td><td style=""border-style:none""><img style=""border:0"" alt=""captcha"" src=""" & C_DIRECTORY_QUICKERSITE & "/asp/includes/captcha.asp"" border=""0"" /></td><td style=""border-style:none""><input style=""width:63px"" type=""text"" value="""" size=""6"" maxlenght=""4"" name=""captcha"" required /></td></tr></table></td></tr><tr><td style=""border-style:none"" align=""left"">"
end if
buildPost=buildPost&"<input class=""art-button"" type='submit' onclick='javascript:this.disabled=true;replyForm"&encrypt(iId)&".submit();' name='dummy' value=" & """" & l("reply") & """" &" />&nbsp;"
buildPost=buildPost&"</td>"
buildPost=buildPost&"</tr>"
buildPost=buildPost&"</table>"
buildPost=buildPost&"</form>"
buildPost=buildPost&"</div>"
end if
dim replyCounter,buildReply
replyCounter=1
if cReplies.Count>0 then
buildPost=buildPost&"<div id='replies"&encrypt(iId)&"' style='display:"&styleReplies&";text-align:left;font-size:0.9em'><a name=""comments""></a><ul>"
for each replyKey in cReplies
buildReply=""
buildReply=buildReply& "<li style='margin-top:15px'>"
buildReply=buildReply&"<a name='"&encrypt(replyKey)&"'></a><strong>" & cReplies(replyKey).contact.sClickNickName & " "& l("says") & ":</strong>&nbsp;"
buildReply=buildReply& "<span style='font-size:0.9em'>(" & formatTimeStamp(cReplies(replyKey).dUpdatedTS) & ")</span><div style='margin-top:5px'>"
if customer.bUseAvatars then
buildReply=buildReply& "<div style=""float:left;margin-top:3px;margin-right:13px;margin-bottom:11px;width:" & customer.iAvatarSize & "px"">"
if convertGetal(logon.contact.iId)=convertGetal(cReplies(replyKey).contact.iId) and convertGetal(logon.contact.iId)<>0 then
buildReply=buildReply& cReplies(replyKey).contact.getClickAvatar()
else
buildReply=buildReply& cReplies(replyKey).contact.getAvatar() 
end if
buildReply=buildReply&"</div>"
end if
buildReply=buildReply& "" & addSmilies(linkUrls(sanitize(cReplies(replyKey).sBody)))
buildReply=buildReply& sFileLink(cReplies(replyKey).sFilename,cReplies(replyKey).sFileDesc,cReplies(replyKey).iFileSize)
if convertGetal(logon.contact.iId)=convertGetal(themeObj.iContactID) and convertGetal(logon.contact.iId)<>0 then
buildReply=buildReply&"<div style=""clear:both""></div>"
buildReply=buildReply& "<div style=""margin-top:5px""><a href="& """" & "default.asp?iId="& encrypt(selectedPage.iId)
buildReply=buildReply& "&amp;iPostID="&encrypt(replyKey)
buildReply=buildReply& "&amp;item="&request("item")
buildReply=buildReply& "&amp;themeAction="&cModReply
buildReply=buildReply& "&amp;" & QS_secCodeURL
buildReply=buildReply& "&amp;startpage="&request("startpage") & """" & " "
buildReply=buildReply& ">" & l("modify") & "</a>"
buildReply=buildReply& " - <a href="& """" & "default.asp?iId="& encrypt(selectedPage.iId)
buildReply=buildReply& "&amp;iPostID="&encrypt(replyKey)
buildReply=buildReply& "&amp;item="&request("item")
buildReply=buildReply& "&amp;btnaction="&cRemoveReply
buildReply=buildReply& "&amp;" & QS_secCodeURL
buildReply=buildReply& "&amp;startpage="&request("startpage") & """" & " "
buildReply=buildReply& "onclick=" & """" & "javascript:return(confirm('"&l("areyousure")&"'));" & """"
buildReply=buildReply& ">" & l("delete") & "</a>"
if not isLeeg(cReplies(replyKey).sFileName) then
buildReply=buildReply& " - <a href="& """" & "default.asp?iId="& encrypt(selectedPage.iId)
buildReply=buildReply& "&amp;iPostID="&encrypt(replyKey)
buildReply=buildReply& "&amp;item="&request("item")
buildReply=buildReply& "&amp;btnaction=removereplyatt"
buildReply=buildReply& "&amp;" & QS_secCodeURL
buildReply=buildReply& "&amp;startpage="&request("startpage") & """" & " "
buildReply=buildReply& "onclick=" & """" & "javascript:return(confirm('"&l("areyousure")&"'));" & """"
buildReply=buildReply& ">" & l("delete") & " attachment</a>"
end if
if convertBool(cReplies(replyKey).bNeedsToBeValidated) then
buildReply=buildReply& " - <a href="& """" & "default.asp?iId="& encrypt(selectedPage.iId)
buildReply=buildReply& "&amp;iPostID="&encrypt(replyKey)
buildReply=buildReply& "&amp;item="&request("item")
buildReply=buildReply& "&amp;btnaction="&cValidateReply
buildReply=buildReply& "&amp;" & QS_secCodeURL
buildReply=buildReply& "&amp;startpage="&request("startpage") & """" & " "
buildReply=buildReply& "onclick=" & """" & "javascript:return(confirm('"&l("areyousure")&"'));" & """"
buildReply=buildReply& ">" & l("validate") & "</a>"
end if
buildReply=buildReply& "</div>"
end if
buildReply=buildReply&"</div><div style=""clear:both""></div>"
buildReply=buildReply&"</li>"
buildPost=buildPost&buildReply
replyCounter=replyCounter+1
next
buildPost=buildPost&"</ul></div>"
end if
set cReplies=nothing
end if
buildPost=buildPost&"</div>"
'tonen of niet?
'if showCompactMode then
buildPost=buildPost&"</div>"
'end if
buildPost=buildPost&"</div>"
if themeObj.postCount=0 then
themeObj.postCount=1
else
themeObj.postCount=0
end if
end function
public function replies(loggedOnAsModerator)
set replies=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, replie, sql
sql="select iId from tblPost where iPostID="& iId
if not loggedOnAsModerator then
sql=sql & " and (bNeedsToBeValidated is null or bNeedsToBeValidated=" & getSQLBoolean(loggedOnAsModerator) & ") "
end if
sql=sql & " order by dCreatedTS asc"
set rs=db.execute(sql)
while not rs.eof 
set replie=new cls_post
replie.pick(rs(0))
replies.Add replie.iID,replie
set replie=nothing
rs.movenext
wend
set rs=nothing
end if
end function
private function getAvatarMail
getAvatarMail=""
if not customer.bUseAvatars then exit function
if isLeeg(logon.contact.sAvatar) then exit function
'if logon.contact.sNickname<>"Pieter" and logon.contact.sNickname<>"Quicker" then exit function
getAvatarMail=logon.contact.sImgTagAvatar(customer.iAvatarSize)
getAvatarMail=replace(getAvatarMail,"margin:0px","margin-right:25px;margin-bottom:25px",1,-1,1)
getAvatarMail=replace(getAvatarMail,"src=","align=""left"" src=",1,-1,1)
end function
public function preparenotifyReply(sReply,sLink)
preparenotifyReply=theme.sBodyNotification
if not isLeeg(sAnName) then
preparenotifyReply=replace(preparenotifyReply,"[QS_theme:replyer]",quotrep(sAnName),1,-1,1)
else
preparenotifyReply=replace(preparenotifyReply,"[QS_theme:replyer]",quotrep(logon.contact.sNickname),1,-1,1)
end if
preparenotifyReply=replace(preparenotifyReply,"[QS_theme:postsubject]",quotrep(parentTopic.sSubject),1,-1,1)
preparenotifyReply=replace(preparenotifyReply,"[QS_theme:reply]","<p>" & getAvatarMail & quotrep(sReply) & "</p>" & sLink,1,-1,1)
preparenotifyReply=replace(preparenotifyReply,"[QS_theme:postlink]",prepareLink,1,-1,1)
preparenotifyReply=replace(preparenotifyReply,vbcrlf,"<br />",1,-1,1)
preparenotifyReply=replace(preparenotifyReply,"<p><br />","<p>",1,-1,1)
preparenotifyReply=replace(preparenotifyReply,"<br /><p>","<p>",1,-1,1)
preparenotifyReply=replace(preparenotifyReply,"<i><i>","<i>",1,-1,1)
preparenotifyReply=replace(preparenotifyReply,"</i></i>","</i>",1,-1,1)
preparenotifyReply=addSmilies(preparenotifyReply)
if bNeedsToBeValidated then
preparenotifyReply=preparenotifyReply & "<p><b>" & l("needsvalidationmail") & "</b></p>"
end if
end function
public function preparenotifyTopic(sReply,sLink)
preparenotifyTopic=theme.sTopicBodyNotification
preparenotifyTopic=replace(preparenotifyTopic,"[QS_theme:name]",quotrep(theme.sName),1,-1,1)
if not isLeeg(sAnName) then
preparenotifyTopic=replace(preparenotifyTopic,"[QS_theme:poster]",quotrep(sAnName),1,-1,1)
else
preparenotifyTopic=replace(preparenotifyTopic,"[QS_theme:poster]",quotrep(logon.contact.sNickname),1,-1,1)
end if
preparenotifyTopic=replace(preparenotifyTopic,"[QS_theme:postsubject]",quotrep(sSubject),1,-1,1)
if theme.bAllowHTML then
preparenotifyTopic=replace(preparenotifyTopic,"[QS_theme:post]",filterJS(sReply) & sLink,1,-1,1)
else
preparenotifyTopic=replace(preparenotifyTopic,"[QS_theme:post]","<p>" & getAvatarMail & quotrep(sReply) & "</p>" & sLink,1,-1,1)
preparenotifyTopic=replace(preparenotifyTopic,vbcrlf,"<br />",1,-1,1)
preparenotifyTopic=replace(preparenotifyTopic,"<p><br />","<p>",1,-1,1)
preparenotifyTopic=replace(preparenotifyTopic,"<br /><p>","<p>",1,-1,1)
preparenotifyTopic=replace(preparenotifyTopic,"<i><i>","<i>",1,-1,1)
preparenotifyTopic=replace(preparenotifyTopic,"</i></i>","</i>",1,-1,1)
end if
preparenotifyTopic=replace(preparenotifyTopic,"[QS_theme:postlink]",prepareLink,1,-1,1)
preparenotifyTopic=addSmilies(preparenotifyTopic)
if bNeedsToBeValidated then
preparenotifyTopic=preparenotifyTopic & "<p><b>" & l("needsvalidationmail") & "</b></p>"
end if
end function
private function prepareLink()
dim iForceOpenID, shortCutReply
if convertGetal(parentTopic.iId)<>0 then
iForceOpenID=encrypt(parentTopic.iId)
shortCutReply="#" & encrypt(iId)
else
iForceOpenID=encrypt(iId)
end if
prepareLink=customer.sQSUrl & "/default.asp?item="& request("item") &"&amp;iId="&encrypt(selectedPage.iId)&"&amp;iPostID="&iForceOpenID & shortCutReply
end function
public function parentTopic
if p_parentTopic is nothing then
set p_parentTopic=new cls_post
p_parentTopic.pick(iPostID)
end if
set parentTopic=p_parentTopic
end function
public function validate()
if convertBool(bNeedsToBeValidated) then
bNeedsToBeValidated=false
save()
end if
end function
Public function modReplyForm
if convertGetal(theme.iContactID)<>convertGetal(logon.contact.iId) or convertGetal(logon.contact.iId)=0 or convertGetal(iId)=0  then exit function
dim postBack
postBack=convertBool(Request.Form ("postBack"))
if postBack then
select case Request.form("modAction")
case l("save")
if convertBool(request.form("bRemoveATT")) then
removeATT
end if
sBody=Request.Form ("sBody")
sAnName=request.form("sAnName")
if save() then Response.Redirect ("default.asp?iId=" & Request.Form ("iId") & "&iPostID=" & encrypt(iPostID) & "&fbMessage=fb_saveOK")
case l("back")
Response.Redirect ("default.asp?iId=" & Request.Form ("iId") & "&iPostID=" & encrypt(iPostID))
end select
end if
modReplyForm="<form action='default.asp' method='post' id=form1 name=form1>"
modReplyForm=modReplyForm&"<input type='hidden' name='themeAction' value='"&cModReply&"' />"
modReplyForm=modReplyForm&"<input type='hidden' name='postBack' value='"&true&"' />"
modReplyForm=modReplyForm&"<input type='hidden' name='iId' value='"&encrypt(selectedpage.iId)&"' />"
modReplyForm=modReplyForm&"<input type='hidden' name='item' value="""&quotrep(request("item"))&""" />"
modReplyForm=modReplyForm&"<input type='hidden' name='iPostID' value='"&encrypt(iId)&"' />"
if convertbool(thisTheme.bAllowAP) and not isLeeg(sAnName) then
modReplyForm=modReplyForm& quotrep(thisTheme.sLabelYourName) & "<br /><input size=""34"" type=""text"" name=""sAnName"" value=""" & sanitize(sAnName) & """ />"
end if
modReplyForm=modReplyForm&"<textarea style='width:100%;margin-bottom:8px' name='sBody' rows=10>" & sanitize(sBody) & "</textarea><br />"
if not isLeeg(sFileName) then
modReplyForm=modReplyForm&"<input type=""checkbox"" value=""1"" name=bRemoveATT /> Delete attachement<br />"
end if
modReplyForm=modReplyForm&"<input class=""art-button"" type='submit' name='modAction' value="""&l("save")&""" /> "
modReplyForm=modReplyForm&"<input class=""art-button"" type='submit' name='modAction' value="""&l("back")&""" />"
modReplyForm=modReplyForm&"</form>"
end function
private function thisTheme
set thisTheme=new cls_theme
thisTheme.pick(iThemeID)
end function
private property get sFileLink(link,desc,size)
sFileLink=""
if not isLeeg(link) then
sFileLink="<div style=""margin-top:6px;margin-bottom:5px"">"
sFileLink=sFileLink & showIcon(GetFileExtension(link))
sFileLink=sFileLink&"<a target=""_blank"" href="""& C_VIRT_DIR &  Application("QS_CMS_userfiles") & "userfiles"  &"/" & link &""">"
if not isLeeg(desc) then
sFileLink=sFileLink&quotrep(desc)
else
sFileLink=sFileLink&"Download attachment"
end if
if convertGetal(size)<>0 then
if convertGetal(size)<1024 then
sFileLink=sFileLink&"&nbsp;(" & size & " kB)"
else
sFileLink=sFileLink&"&nbsp;(" & round(size/1024,2) & " MB)"
end if
end if
sFileLink=sFileLink&"</a>"
sFileLink=sFileLink& "</div>"
if lcase(right(link,3))="jpg" then
sFileLink=sFileLink& "<div>"
sFileLink=sFileLink& "<a target=""_blank"" title=""Click to enlarge"" href=""" & C_VIRT_DIR &  Application("QS_CMS_userfiles") & "userfiles"  &"/" & link & """>"
sFileLink=sFileLink& "<img style=""box-shadow:1px 1px 3px #CCC;border-style:none"" src=""" & C_DIRECTORY_QUICKERSITE & "/showthumb.aspx?maxsize=350&amp;img=" & C_VIRT_DIR &  Application("QS_CMS_userfiles") & "userfiles"  &"/" & link & """ alt=""Click to enlarge"" />"
sFileLink=sFileLink& "</a></div>"
end if
end if
end property
public function removeATT
if not isLeeg(sFilename) then
dim fso
set fso=server.CreateObject ("scripting.filesystemobject")
if fso.FileExists (server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/" & sFilename)) then
fso.DeleteFile server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/" & sFilename)
end if
set fso=nothing
bRemoveAnyway=true
sFilename=""
sFileDesc=""
iFileSize=0
save()
end if
end function
end class%>
