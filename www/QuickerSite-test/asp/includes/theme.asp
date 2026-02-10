
<%class cls_theme
Public iId, sName, iPageSize, bOnline, bAllowComments, bAllowHTML, bPushRSS, startpage, totalNumber, sNextPage, sPreviousPage, bLocked, iType, bAllowAP
Public iContactID, iLimitReplyTo, iLimitTopicTo, sBodyNotification, sSubjectNotification, bForwardPostsToModerator, iSubLevel, bSmileys, bCompactList
Public sTopicBodyNotification, sTopicSubjectNotification, isForcedOpen, postCount, sColorUneven, sColorEven, iWidth, sCode, recomWidth, bUpload,bValidation
Public iSearchType, bEmbed, sLabelYourName, bFileUploads
Private p_allowNewPost
Private Sub Class_Initialize
On Error Resume Next
postCount=0
iLimitReplyTo=500
iWidth=500
recomWidth=iWidth
iLimitTopicTo=2500
iPageSize=20
bOnline=true
bEmbed=false
bLocked=false
bAllowComments=true
bAllowHTML=false
bPushRSS=true
iType=QS_theme_cd
iContactID=null
isForcedOpen=false
bSmileys=true
bUpload=false
bCompactList=false
bValidation=false
bAllowAP=false
bFileUploads=false
iSearchType=1
sLabelYourName="(your name)"
bForwardPostsToModerator=true
sBodyNotification="<p>You are receiving this email because you requested to be notified when somebody responded to the topic <em>[QS_theme:postsubject]</em>.</p><p><strong>[QS_theme:Replyer]</strong> has replied.</p><p><i>[QS_theme:reply]</i></p><p>Click the link below or paste it into your browser to read the reply online.</p><p><a href='[QS_theme:postlink]'>[QS_theme:postlink]</a></p>"
sSubjectNotification="Notification of new Reply"
sTopicBodyNotification="<p>You are receiving this email because you requested to be notified when somebody posted a new topic in the theme <em>[QS_theme:Name]</em>.</p><p><strong>[QS_theme:Poster]</strong> has posted the topic <em>[QS_theme:postsubject]</em>.</p><p><i>[QS_theme:post]</i></p><p>Click the link below or paste it into your browser to read the topic online.</p><p><a href='[QS_theme:postlink]'>[QS_theme:postlink]</a></p>"
sTopicSubjectNotification="Notification of new Topic"
iSubLevel=QS_theme_sublevel_topic
if isNumeriek(Request.Querystring ("startpage")) then
startpage=convertGetal(Request.Querystring ("startpage"))
else
startpage=convertGetal(Request("startpage"))
end if
if startpage=0 or startpage>500 then
startpage=1
end if
pick(decrypt(request("iThemeID")))
On Error Goto 0
end sub
private property get incJS
incJS=vbcrlf& "<script type=""text/javascript"">" & vbcrlf
incJS=incJS & "function compact(id){" & vbcrlf
incJS=incJS & "if(document.getElementById('topicTable'+id).style.display=='block'){document.getElementById('topicTable'+id).style.display='none';document.getElementById('zoom'+id).src='" & C_DIRECTORY_QUICKERSITE & "/fixedImages/plus2.gif';return false;}else{document.getElementById('topicTable'+id).style.display='block';document.getElementById('zoom'+id).src='" & C_DIRECTORY_QUICKERSITE & "/fixedImages/minus2.gif'};return false;}" & vbcrlf
incJS=incJS & "function uncompact(id){" & vbcrlf
incJS=incJS & "if(document.getElementById('topicTable'+id).style.display=='none'){document.getElementById('topicTable'+id).style.display='block';document.getElementById('zoom'+id).src='" & C_DIRECTORY_QUICKERSITE & "/fixedImages/minus2.gif';return false;}else{document.getElementById('topicTable'+id).style.display='none';document.getElementById('zoom'+id).src='" & C_DIRECTORY_QUICKERSITE & "/fixedImages/plus2.gif'};return false;}"
incJS=incJS & "</script>"& vbcrlf
end property
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblTheme where iCustomerID="&cid&" and iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sName	= rs("sName")
iPageSize	= rs("iPageSize")
bOnline	= rs("bOnline")
bLocked	= rs("bLocked")
bAllowComments	= rs("bAllowComments")
bAllowHTML	= rs("bAllowHTML")
bPushRSS	= rs("bPushRSS")
iType	= rs("iType")
iContactID	= rs("iContactID")
iLimitReplyTo	= rs("iLimitReplyTo")
iLimitTopicTo	= rs("iLimitTopicTo")
sBodyNotification	= rs("sBodyNotification")
sSubjectNotification	= rs("sSubjectNotification")
bForwardPostsToModerator	= rs("bForwardPostsToModerator")
iSubLevel	= rs("iSubLevel")
sTopicBodyNotification	= rs("sTopicBodyNotification")
sTopicSubjectNotification	= rs("sTopicSubjectNotification")
sColorUneven	= rs("sColorUneven")
sColorEven	= rs("sColorEven")
iWidth	= rs("iWidth")
sCode	= rs("sCode")
bSmileys	= rs("bSmileys")
bUpload	= rs("bUpload")
bCompactList	= rs("bCompactList")
bValidation	= rs("bValidation")
iSearchType	= rs("iSearchType")
bEmbed	= rs("bEmbed")
bAllowAP	= rs("bAllowAP")
sLabelYourName	= rs("sLabelYourName")
bFileUploads	= rs("bFileUploads")
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
if isLeeg(sBodyNotification) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sSubjectNotification) then
check=false
message.AddError("err_mandatory")
end if
if convertBool(bAllowAP) then
if isLeeg(sLabelYourName) then
check=false
message.AddError("err_mandatory")
end if
end if
if iType=QS_theme_pb then
if isLeeg(iContactID) then
check=false
message.AddError("err_mandatory")
end if
end if
if not isLeeg(sCode) then
dim checkRS
set checkRS=db.execute("select count(iId) from tblTheme where iId<>" & convertGetal(iId) & " and iCustomerID=" & cID & " and sCode='" & ucase(sCode) & "'")
if clng(checkRS(0))>0 then
check=false
message.AddError("err_doublefeed")
end if
set checkRS=nothing
end if
End Function
Public Function Save
if check() then
save=true
else
save=false
exit function
end if
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblTheme where 1=2"
rs.AddNew
else
rs.Open "select * from tblTheme where iId="& iId
end if
rs("sName")	= left(sName,255)
rs("iPageSize")	= iPageSize
rs("bOnline")	= bOnline
rs("bLocked")	= bLocked
rs("bAllowComments")	= bAllowComments
rs("bAllowHTML")	= bAllowHTML
rs("bPushRSS")	= bPushRSS
rs("iCustomerID")	= cId
rs("iType")	= iType
rs("iContactID")	= iContactID
rs("iLimitReplyTo")	= iLimitReplyTo
rs("iLimitTopicTo")	= iLimitTopicTo
rs("sBodyNotification")	= sBodyNotification
rs("sSubjectNotification")	= sSubjectNotification
rs("bForwardPostsToModerator")	= bForwardPostsToModerator
rs("iSubLevel")	= iSubLevel
rs("sTopicBodyNotification")	= sTopicBodyNotification
rs("sTopicSubjectNotification")	= sTopicSubjectNotification
rs("sColorUneven")	= sColorUneven
rs("sColorEven")	= sColorEven
rs("iWidth")	= iWidth
rs("sCode")	= sCode
rs("bSmileys")	= bSmileys
rs("bUpload")	= bUpload
rs("bCompactList")	= bCompactList
rs("bValidation")	= bValidation
rs("iSearchType")	= iSearchType
rs("bEmbed")	= bEmbed
rs("bAllowAP")	= bAllowAP
rs("sLabelYourName")	= sLabelYourName
rs("bFileUploads")	= bFileUploads
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
'treat subsriptions
dim rsDelSub, rsPosts
if iSubLevel=QS_theme_sublevel_none then
set rsPosts=db.execute("select iId from tblPost where iThemeID="& iId)
while not rsPosts.eof
set rsDelSub=db.execute("delete from tblThemeTopicSubscription where iPostId="& rsPosts(0))
set rsDelSub=nothing
rsPosts.movenext
wend
set rsPosts=nothing
end if
if iSubLevel=QS_theme_sublevel_authortopic then
set rsPosts=db.execute("select iId, iContactID from tblPost where iThemeID="& iId)
while not rsPosts.eof
set rsDelSub=db.execute("delete from tblThemeTopicSubscription where iPostId="& rsPosts(0) & " and iContactID<>"& rsPosts(1))
set rsDelSub=nothing
rsPosts.movenext
wend
set rsPosts=nothing
end if
if iSubLevel=QS_theme_sublevel_topic or iSubLevel=QS_theme_sublevel_none or iSubLevel=QS_theme_sublevel_authortopic then
set rsDelSub=db.execute("delete from tblThemeSubscription where iThemeID="& iId)
set rsDelSub=nothing
end if
if not bAllowComments then
'alle commentaren eruit zwieren...
set rsDelSub=db.execute("delete from tblPost where iThemeID="& iId & " and (iPostID is not null)")
set rsDelSub=nothing
end if
'refresh RSS content
clearRSScache
'refresh cachethemes
customer.cacheThemes()
end function
public sub clearRSScache
dim rs
set rs=db.execute("select iId from tblPage where iThemeID="& iId)
dim page
while not rs.eof
set page=new cls_page
page.pick(rs(0))
page.clearRSScache()
set page=nothing
rs.movenext
wend
set rs=nothing
end sub
public function getRequestValues()
sName	= convertStr(Request.Form("sName"))
iPageSize	= convertGetal(Request.Form("iPageSize"))
bOnline	= convertBool(Request.Form("bOnline"))
bLocked	= convertBool(Request.Form ("bLocked"))
bAllowComments	= convertBool(Request.Form("bAllowComments"))
bAllowHTML	= convertBool(Request.Form("bAllowHTML"))
bPushRSS	= convertBool(Request.Form("bPushRSS"))
iType	= convertGetal(Request.Form ("iType"))
iContactID	= convertLng(Request.Form ("iContactID"))
iLimitReplyTo	= convertGetal(Request.Form ("iLimitReplyTo"))
iLimitTopicTo	= convertGetal(Request.Form ("iLimitTopicTo"))
sBodyNotification	= removeEmptyP(Request.Form ("sBodyNotification"))
sSubjectNotification	= convertStr(Request.Form ("sSubjectNotification"))
bForwardPostsToModerator	= convertBool(Request.Form ("bForwardPostsToModerator"))
iSubLevel	= convertGetal(Request.Form ("iSubLevel"))
sTopicBodyNotification	= convertStr(Request.Form ("sTopicBodyNotification"))
sTopicSubjectNotification	= convertStr(Request.Form ("sTopicSubjectNotification"))
sColorUneven	= convertStr(Request.Form ("sColorUneven"))
sColorEven	= convertStr(Request.Form ("sColorEven"))
iWidth	= convertGetal(Request.Form ("iWidth"))
sCode	= ucase(convertStr(Request.Form ("sCode")))
bSmileys	= convertBool(Request.Form ("bSmileys"))
bUpload	= convertBool(Request.Form ("bUpload"))
bCompactList	= convertBool(Request.Form ("bCompactList"))
bValidation	= convertBool(Request.Form ("bValidation"))
iSearchType	= convertGetal(Request.Form ("iSearchType"))
bEmbed	= convertBool(Request.Form ("bEmbed"))
bAllowAP	= convertBool(Request.Form ("bAllowAP"))
sLabelYourName	= convertStr(request.form("sLabelYourName"))
bFileUploads	= convertBool(Request.Form ("bFileUploads"))
end function
public function remove
if not isLeeg(iId) then
dim cPosts, postKey
set cPosts=posts
for each postKey in cPosts
cPosts(postKey).remove
next 
set cPosts=nothing
dim rs
set rs=db.execute("update tblPage set iThemeID=null where iThemeID="& iId)
set rs=nothing
set rs=db.execute("delete from tblThemeSubscription where iThemeID="& iId)
set rs=nothing
set rs=db.execute("delete from tblTheme where iId="& iId)
set rs=nothing
end if
end function
public property get allowPost
allowPost=true
if bLocked then 
allowPost=false 'theme/blog is gelocked - geen nieuwe posts
elseif convertGetal(logon.contact.iId)=0 then
allowPost=false 'gebruiker niet aangelogd
elseif logon.contact.istatus<cs_write then 
allowPost=false 'indien de aangelogde gebruiker geen schrijfrechten heeft, hier blokkeren
elseif iType=QS_theme_pb and logon.contact.iId<>iContactID then 
allowPost=false 'in geval van personal blog mag niemand anders posten
end if
end property
public property get allowSub
allowSub=true
if bLocked then 
allowSub=false 'theme/blog is gelocked - geen nieuwe subscriptions
elseif convertGetal(logon.contact.iId)=0 then
allowSub=false 'gebruiker niet aangelogd
elseif iSubLevel=QS_theme_sublevel_none then
allowSub=false 'forum accepteerd geen subscriptions
end if
end property
public function build()
build=""
dim post
if bOnline and isNumeriek(iId) then
set post=new cls_post 'verwacht request("iPostID")
'theme-actionhandler
select case Request("themeAction")
case cSubscribeToTheme
checkCSRF()
if subscribeToTheme then
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId)&"&fbMessage=fb_subscribetheme&item=" & Request("item"))
end if
case cUnSubscribeToTheme
checkCSRF()
if unsubscribeFromTheme then
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId)&"&fbMessage=fb_subscriptionthemecancelled&item=" & Request("item"))
end if
case cSubscribeToTopic
checkCSRF()
if subscribeToTopic(post) then
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId) & "&iPostID="&encrypt(post.iId)&"&fbMessage=fb_subscribetopic&item=" & Request("item"))
end if
case cUnSubscribeToTopic
checkCSRF()
if unsubscribeFromTopic(post) then
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId) & "&iPostID="&encrypt(post.iId)&"&fbMessage=fb_subscriptiontopiccancelled&item=" & Request("item"))
end if
case cPostTopic
checkCSRF()
'de functie postTopic zorgt voor redirectie...
build=build&shortNavBar
build=build&postTopic(post)
case cPostReply
checkCSRF()
'de functie buildPost zorgt voor redirectie
build=build&post.buildPost(me,true)
case cModReply
checkCSRF()
'de functie buildPost zorgt voor redirectie
build=build&post.modReplyForm
case cSearchTheme
build=build & buildSearch
case clSubs
checkCSRF()
if logon.contact.iId=iContactID then 
build=build & buildSubs
end if
case cdSub
checkCSRF()
if logon.contact.iId=iContactID then 
'remove themesub
db.execute("delete from tblThemeSubscription where iContactID=" & convertGetal(decrypt(Request.QueryString ("iContactID"))) & " and iThemeID=" & convertGetal(decrypt(Request.QueryString ("iThemeID"))))
Response.Redirect ("default.asp?themeAction=" & clSubs & "&iId=" & encrypt(selectedPage.iID))
end if
case else
if convertGetal(post.iId)<>0 then
'de topic tonen, inclusief de replies
build=build&shortNavBar 'kort navigatie terugn naar lijst
build=build&post.buildPost(me,true)
build=build&shortNavBar 'ook onderaan
selectedPage.sSEOtitle=quotrep(post.sSubject & " | "  & customer.siteTitle)
else
'only show paged topics
dim cposts, postKey
set cposts=posts
dim cnavbar
cnavbar=navbar(cposts.count)
build=build&cnavbar
for each postKey in cposts
build=build & cposts(postKey).buildPost(me,false)
next
if cposts.count>0 then
build=build&cnavbar
end if
end if
end select
build=incJS&replace(build,"[","|R|R|R|",1,-1,1)
end if
end function
public function buildSearch
buildSearch=buildSearch	&"<form method='post' name='buildSearch' action='default.asp'>"
buildSearch=buildSearch	&QS_secCodeHidden
buildSearch=buildSearch	&"<input type='hidden' name='iId' value='"&encrypt(selectedPage.iId)&"' />"
buildSearch=buildSearch	&"<input type='hidden' name='themeAction' value='"&cSearchTheme&"' />"
buildSearch=buildSearch	&"<table align='center' cellpadding=5 border=0>"
buildSearch=buildSearch	& "<tr>"
buildSearch=buildSearch	& "<td width='50%' align='right' valign='middle'><input required type=""text"" maxlength=""60"" id='sValueThemeSearch' name='sValueThemeSearch' value="""& quotrep(request.form("sValueThemeSearch")) &""" /></td>"
buildSearch=buildSearch	& "<td width='50%' align='left' valign='middle'><input class=""art-button"" type='submit' name='dummy' value="""&l("search")&""" /></td>"
'document.getElementById('theInput').focus(); 
buildSearch=buildSearch	& "</tr>"
dim sValueThemeSearch
sValueThemeSearch=Request("sValueThemeSearch")
if not isLeeg(sValueThemeSearch) then
dim copyPosts
set copyPosts=searchPosts(sValueThemeSearch)
buildSearch=buildSearch & "<tr>"
buildSearch=buildSearch & "<td colspan=2  valign='middle' align='center'>" & l("thereAre") & " "& copyPosts.count & " " & l("resultsForSearch") & ": '"& sanitize(treatConstants(sValueThemeSearch,false)) &"'</td>"
buildSearch=buildSearch & "</tr>"
if copyPosts.count>0 then
buildSearch=buildSearch & "<tr><td colspan=2>"
buildSearch=buildSearch & "<ul>"
dim post 
for each post in copyPosts
buildSearch=buildSearch & "<li><a href='default.asp?iId=" & encrypt(selectedPage.iId) & "&amp;iPostID=" & encrypt(post) & "'><b>" & quotrep(copyPosts(post).sSubject) & "</b></a><br /><i>" & replace(left(convertStr(treatConstants(removeHTML(copyPosts(post).sBody),false)),70),sValueThemeSearch,"<b>"& sanitize(sValueThemeSearch) &"</b>",1,-1,1) & "...</i></li>"
next 
buildSearch=buildSearch & "</ul>"
buildSearch=buildSearch & "</td></tr>"
end if
set copyPosts=nothing
end if
buildSearch=buildSearch	& "<tr><td colspan=2 valign='middle' align=center><a href='default.asp?iId=" & encrypt(selectedPage.iId) & "'>" & l("back") & "</a></td></tr></table></form>"
buildSearch=buildSearch	&"<script type=""text/javascript"">document.getElementById('sValueThemeSearch').focus();</script>"
end function
public function postTopic(postObj)
if not allowPost then Response.Redirect ("default.asp?iId="&encrypt(selectedPage.iId)) 'basis beveiliging
'extra beveiliging
if convertGetal(postObj.iId)<>0 and logon.contact.iId<>postObj.iContactID and iContactID<>logon.contact.iId then 'aanpassen voor moderator
Response.Redirect ("default.asp?iId="&encrypt(selectedPage.iId))
end if
dim postback
postback=convertBool(Request.Form ("postback")) 
if postback then
select case Request.Form ("btnaction")
case l("delete") 
checkCSRF()
if convertGetal(iContactID)=convertGetal(logon.contact.iId) then 
postObj.remove()
Response.Redirect ("default.asp?iId="&encrypt(selectedPage.iId)&"&fbMessage=fb_topicremoved&item="& request("item"))
else
Response.Redirect ("default.asp?iId="&encrypt(selectedPage.iId))
end if
case l("validate")
checkCSRF()
postObj.sSubject	= left(Request.Form ("sSubject"),100)
if bAllowHTML then
if convertBool(bEmbed) then
postObj.sBody	= removeEmptyP(Request.Form ("sBody"))
else
'extra filtering van JS
postObj.sBody	= removeEmptyP(removeJS(Request.Form ("sBody")))
end if
else
postObj.sBody	= left(Request.Form ("sBody"),(iLimitTopicTo+50))
end if
postObj.iThemeID	= iId
if convertGetal(request.form("iChangeThemeID"))<>0 then 
postObj.iThemeID=convertGetal(request.form("iChangeThemeID"))
end if
if convertGetal(iContactID)=convertGetal(logon.contact.iId) then 
postObj.validate()
if convertBool(request.form("bRemoveATT")) then
postObj.removeATT()
'response.end 
end if
Response.Redirect ("default.asp?iId="&encrypt(selectedPage.iId)&"&fbMessage=fb_topicvalidated&item="& request("item"))
else
Response.Redirect ("default.asp?iId="&encrypt(selectedPage.iId))
end if
case else
checkCSRF()
postObj.sSubject	= left(Request.Form ("sSubject"),100)
if bFileUploads and convertBool(logon.authenticatedIntranet) then
postObj.sFileName	= left(Request.Form ("sFileName"),50)
postObj.sFileDesc	= left(Request.Form ("sFileDesc"),50)
postObj.iFileSize	= left(Request.Form ("iFileSize"),50)
end if
if bAllowHTML then
if convertBool(bEmbed) then
postObj.sBody	= removeEmptyP(Request.Form ("sBody"))
else
'extra filtering van JS
postObj.sBody	= removeEmptyP(removeJS(Request.Form ("sBody")))
end if
else
postObj.sBody	= left(Request.Form ("sBody"),(iLimitTopicTo+50))
end if
postObj.iThemeID	= iId
if convertGetal(request.form("iChangeThemeID"))<>0 then 
postObj.iThemeID=convertGetal(request.form("iChangeThemeID"))
end if
If LCase(session("CAPTCHA")) <> LCase(Left(Request.Form("CAPTCHA"),4)) Then
message.AddError("err_captcha")
elseif postObj.save() then
if convertBool(request.form("bRemoveATT")) then
postObj.removeATT()
'response.end 
end if
dim feedbackMessages
feedbackMessages="fb_topicsave"
if postObj.bNeedsToBeValidated then
feedbackMessages=feedbackMessages& ",fb_explvalidation"
end if
if convertBool(Request.Form ("bSubscribe")) then
if subscribeToTopic(postObj) then
if postObj.bNeedsToBeValidated then
Response.Redirect ("default.asp?iId="&encrypt(selectedPage.iId)&"&fbMessage="&feedbackMessages&"&item=" & Request("item"))
else
Response.Redirect ("default.asp?iId="&encrypt(selectedPage.iId)&"&iPostID="&encrypt(postObj.iID)&"&fbMessage="&feedbackMessages&"&item=" & Request("item"))
end if
end if
else
if postObj.bNeedsToBeValidated then
Response.Redirect ("default.asp?iId="&encrypt(selectedPage.iId)&"&fbMessage="&feedbackMessages&"&item=" & Request("item"))
else
Response.Redirect ("default.asp?iId="&encrypt(selectedPage.iId)&"&iPostID="&encrypt(postObj.iID)&"&fbMessage="&feedbackMessages&"&item=" & Request("item"))
end if
end if
end if 
end select
end if
postTopic=postTopic&"<form method='post' name='postTopic' action='default.asp'>"
postTopic=postTopic&QS_secCodeHidden
postTopic=postTopic&"<input type='hidden' name='iId' value='"&encrypt(selectedPage.iId)&"' />"
postTopic=postTopic&"<input type='hidden' name='themeAction' value='"&cPostTopic&"' />"
postTopic=postTopic&"<input type='hidden' name='iPostID' value='"&encrypt(postObj.iId)&"' />"
postTopic=postTopic&"<input type='hidden' name='item' value="""&quotrep(request("item"))&""" />"
postTopic=postTopic&"<input type='hidden' name='postBack' value="""&true&""" />"
if bFileUploads then
'response.write "FN: " & postObj.sFileName
postTopic=postTopic&"<input type='hidden' name=""sFileName"" value=""" & quotrep(postObj.sFileName) & """ />"
postTopic=postTopic&"<input type='hidden' name=""sFileDesc"" value=""" & quotrep(postObj.sFileDesc) & """ />"
postTopic=postTopic&"<input type='hidden' name=""iFileSize"" value=""" & quotrep(postObj.iFileSize) & """ />"
end if
postTopic=postTopic&"<table align=center>"
postTopic=postTopic& "<tr>"
postTopic=postTopic& "<td class='QSlabel'>"&l("subject")&":*</td>"
postTopic=postTopic& "<td><input required type='text' style='width:100%' size='60' maxlength='100' name='sSubject' value=" & """" & sanitize(postObj.sSubject) & """" &" /></td>"
postTopic=postTopic& "</tr>"
postTopic=postTopic& "<tr>"
postTopic=postTopic& "<td class='QSlabel'>"&l("text")&":*</td>"
bThemeEmbed=convertBool(bEmbed)
if bAllowHTML then
postTopic=postTopic& "<td><input type='hidden' name='superflus' value='dsfqsdfqsdfqsdf' />" 
if bUpload then 
'create a subfolder if needed.
logon.contact.createUserFilesFolder()
postTopic=postTopic&  dumpFCKInstance(postObj.sBody,"siteBuilderMailNoSourceUpload","sBody") 
else
postTopic=postTopic&  dumpFCKInstance(postObj.sBody,"siteBuilderMailNoSource","sBody") 
end if
if QS_EDITOR=1 then
postTopic=postTopic & "<a href='#' onclick=" & """" & "javascript:resizeiframe('sBody___Frame');return false;"& """" & " style='font-size:0.8em'>" & l("enlargeeditor")&"</a>"
end if
if convertBool(bFileUploads) and convertBool(logon.authenticatedIntranet) then
if request.form("sFileName")<>"" then
if request.form("sFileDesc")<>"" then
postTopic=postTopic& "<div style=""clear:both""></div><span id=""sFileD"">The file <i><b>" & quotrep(request.form("sFileDesc")) & "</b></i> is attached to this topic</span>"
else
postTopic=postTopic& "<div style=""clear:both""></div><span id=""sFileD"">The file <i><b>" & quotrep(request.form("sFileName")) & "</b></i> is attached to this topic</span>"
end if
elseif not isLeeg(postObj.sFilename) then
postTopic=postTopic& "<div style=""clear:both""></div><span id=""sFileD"">The file <i><b>"
if isLeeg(postObj.sFileDesc) then
postTopic=postTopic&postObj.sFilename
else
postTopic=postTopic&quotrep(postObj.sFileDesc)
end if
postTopic=postTopic&"</b></i> is attached to this topic. - <a href=""default.asp?pageAction=fileupload&amp;iPostID=" & encrypt(postObj.iId) & "&amp;iThemeID=" & encrypt(iId) & """ class=""QSPPAVATAR"">Replace</a></span>"
else
postTopic=postTopic& "<div style=""clear:both""></div><span id=""sFileD""><a href=""default.asp?pageAction=fileupload&amp;iPostID=" & encrypt(postObj.iId) & "&amp;iThemeID=" & encrypt(iId) & """ class=""QSPPAVATAR"">Attach file</a></span>"
end if
if not isLeeg(postObj.sFilename) then
postTopic=postTopic& "<div><input type=""checkbox"" name=""bRemoveATT"" value=""1"" /> Delete attachment</div>"
end if
end if
postTopic=postTopic & "</td>"
else
postTopic=postTopic& "<td><textarea style='width:100%' rows='15' name='sBody' onKeyDown=""javascript:textCounter(this.form.sBody,'remLen',"&iLimitTopicTo&");"" onKeyUp=""javascript:textCounter(this.form.sBody,'remLen',"&iLimitTopicTo&");"">" & sanitize(postObj.sBody) & "</textarea>"
if convertBool(bFileUploads) and convertBool(logon.authenticatedIntranet) then
if request.form("sFileName")<>"" then
if request.form("sFileDesc")<>"" then
postTopic=postTopic& "<div style=""clear:both""></div><span id=""sFileD"">The file <i><b>" & quotrep(request.form("sFileDesc")) & "</b></i> is attached to this topic</span>"
else
postTopic=postTopic& "<div style=""clear:both""></div><span id=""sFileD"">The file <i><b>" & quotrep(request.form("sFileName")) & "</b></i> is attached to this topic</span>"
end if
elseif not isLeeg(postObj.sFilename) then
postTopic=postTopic& "<div style=""clear:both""></div><span id=""sFileD"">The file <i><b>"
if isLeeg(postObj.sFileDesc) then
postTopic=postTopic&postObj.sFilename
else
postTopic=postTopic&quotrep(postObj.sFileDesc)
end if
postTopic=postTopic&"</b></i> is attached to this topic. - <a href=""default.asp?pageAction=fileupload&amp;iPostID=" & encrypt(postObj.iId) & "&amp;iThemeID=" & encrypt(iId) & """ class=""QSPPAVATAR"">Replace</a></span>"
else
postTopic=postTopic& "<div style=""clear:both""></div><span id=""sFileD""><a href=""default.asp?pageAction=fileupload&amp;iPostID=" & encrypt(postObj.iId) & "&amp;iThemeID=" & encrypt(iId) & """ class=""QSPPAVATAR"">Attach file</a></span>"
end if
if not isLeeg(postObj.sFilename) then
postTopic=postTopic& "<div><input type=""checkbox"" name=""bRemoveATT"" value=""1"" /> Delete attachment</div>"
end if
end if
if convertBool(bSmileys) then postTopic=postTopic&"<div style=""clear:both""></div><div style='margin:2px'>" & shortListOfSmilies("postTopic.sBody")  & "</div>"
postTopic=postTopic& "</td>"
end if
postTopic=postTopic& "</tr>"
'change forumID
if convertGetal(postObj.iId)<>0 then
postTopic=postTopic & "<tr><td class='QSlabel'>&nbsp;</td><td><select name=""iChangeThemeID"">" 
dim cThemes,cTheme
set cThemes=customer.themes
for each cTheme in cThemes
if convertGetal(iType)=convertGetal(cThemes(cTheme).iType) then
postTopic=postTopic & "<option value=""" & cTheme & """ "
if convertGetal(request.form("iChangeThemeID"))=convertGetal(cTheme) then
postTopic=postTopic & " selected=""selected"" "
elseif convertGetal(iId)=convertGetal(cTheme) and convertGetal(request.form("iChangeThemeID"))=0 then
postTopic=postTopic & " selected=""selected"" "
end if
postTopic=postTopic & ">" & quotrep(cThemes(cTheme).sName) & "</option>"
end if
next
set cThemes=nothing
postTopic=postTopic & "</select></td><tr>"
end if
postTopic=postTopic& "<tr>"
postTopic=postTopic& "<td class='QSlabel'>&nbsp;</td>"
postTopic=postTopic& "<td><img src='" & C_DIRECTORY_QUICKERSITE & "/asp/includes/captcha.asp' alt='' style='vertical-align: middle;border-style:none;margin:0px 6px 0px 0px' /><input style='width:63px' type='text' size='6' maxlength='4' name='captcha' required />&nbsp;" & l("captcha") &"</td>"
postTopic=postTopic& "</tr>"
if isLeeg(postObj.iId) and iSubLevel>=QS_theme_sublevel_authortopic and not logon.contact.bHasSubscribedToTheme(iId) then
dim defaultSub
if isLeeg(Request.Form ("btnaction")) then
defaultSub=convertChecked(true)
else
defaultSub=convertChecked(Request.Form ("bSubscribe"))
end if
postTopic=postTopic& "<tr>"
postTopic=postTopic& "<td class='QSlabel'>"&l("notificationreplies")&"</td>"
postTopic=postTopic& "<td><input type='checkbox' value='"&true&"' name='bSubscribe' "& defaultSub &" />&nbsp;"&l("expl_subscribe_topic_short")&"</td>"
postTopic=postTopic& "</tr>"
end if
postTopic=postTopic& "<tr>"
postTopic=postTopic& "<td>&nbsp;</td>"
postTopic=postTopic& "<td>(*) "&l("mandatory")&"</td>"
postTopic=postTopic& "</tr>"
postTopic=postTopic& "<tr>"
postTopic=postTopic& "<td>&nbsp;</td>"
postTopic=postTopic& "<td><input class=""art-button"" type=""submit"" name=""btnaction"" value="& """" & l("save")  & """" & " />"
if convertGetal(postObj.iId)<>0 and iContactID=logon.contact.iId then
postTopic=postTopic& "&nbsp;<input class=""art-button"" onclick=" &"""" & "javascript:return confirm('"&l("areyousure")&"')" & """" & " type='submit' name='btnaction' value="& """" & l("delete")  & """" & " />"
if convertBool(postObj.bNeedsToBeValidated) then
postTopic=postTopic& "&nbsp;<input class=""art-button"" onclick=" &"""" & "javascript:return confirm('"&l("areyousure")&"')" & """" & " type='submit' name='btnaction' value="& """" & l("validate")  & """" & " />"
end if
end if
if not bAllowHTML then
postTopic=postTopic&"&nbsp;<span id='remLen'>" & iLimitTopicTo & "</span>&nbsp;"& l("charactersleft")
end if
postTopic=postTopic& "</td>"
postTopic=postTopic& "</tr>"
postTopic=postTopic& "</table>"
postTopic=postTopic& "</form>"
set postObj=nothing
end function
public function contact
set contact=new cls_contact
contact.pick(iContactID)
end function
private function subscribeToTopic(postObj)
subscribeToTopic=false
if not allowSub then exit function 'algemene check
if iSubLevel<QS_theme_sublevel_authortopic then exit function 'topic-inschrijvingen zijn niet toegelaten
'quick and dirty - to be revised
on error resume next
if iSubLevel=QS_theme_sublevel_authortopic then
if postObj.iContactID<>logon.contact.iId then exit function 'prevent other users to sign in...
end if
if not logon.contact.bHasSubscribedToTheme(iId) then
dim rs
set rs=db.getDynamicRS
rs.Open "select * from tblThemeTopicSubscription where 1=2"
rs.AddNew
rs("iContactID")=convertGetal(logon.contact.iId)
rs("iPostID")=postObj.iId
rs.update()
rs.close
set rs=nothing
if err.number=0 then
subscribeToTopic=true
err.Clear ()
end if
end if
on error goto 0
end function
private function unsubscribeFromTopic(postObj)
unsubscribeFromTopic=false
'quick and dirty - to be revised
on error resume next
dim rs
set rs=db.execute("delete from tblThemeTopicSubscription where iPostID="& postObj.iId & " and iContactID="& convertGetal(logon.contact.iId))
set rs=nothing
if err.number=0 then
unsubscribeFromTopic=true
err.Clear ()
end if
on error goto 0
end function
private function subscribeToTheme()
subscribeToTheme=false
if not allowSub then exit function 'algemene check
if iSubLevel<QS_theme_sublevel_theme then exit function 'thema-inschrijvingen zijn niet toegelaten
'quick and dirty - to be revised
on error resume next
dim rs
set rs=db.getDynamicRS
rs.Open "select * from tblThemeSubscription where 1=2"
rs.AddNew
rs("iContactID")=convertGetal(logon.contact.iId)
rs("iThemeID")=convertGetal(iId)
rs.update()
rs.close
set rs=nothing
'verwijderen van alle topics-inschrijvingegn voor dit theme
dim rsPosts,rsDelSub
set rsPosts=db.execute("select iId from tblPost where iThemeID="& iId)
while not rsPosts.eof
set rsDelSub=db.execute("delete from tblThemeTopicSubscription where iPostId="& rsPosts(0) & " and iContactID=" &convertGetal(logon.contact.iId))
set rsDelSub=nothing
rsPosts.movenext
wend
set rsPosts=nothing
if err.number=0 then
subscribeToTheme=true
err.Clear ()
end if
on error goto 0
end function
private function unsubscribeFromTheme()
unsubscribeFromTheme=false
'quick and dirty - to be revised
on error resume next
dim rs
set rs=db.execute("delete from tblThemeSubscription where iThemeID="&convertGetal(iId) & " and iContactID="& convertGetal(logon.contact.iId))
set rs=nothing
if err.number=0 then
unsubscribeFromTheme=true
err.Clear ()
end if
on error goto 0
end function
public function shortNavBar
shortNavBar="<div style='text-align:center;width:100%;font-size:0.9em;margin-bottom:10px'>"
if allowPost then shortNavBar=shortNavBar& "<a href='default.asp?"&QS_secCodeURL&"&amp;iId="&encrypt(selectedPage.iId)&"&amp;themeAction="&cPostTopic&"&amp;item="& request("item") &"'><b>"&l("newtopic")&"</b></a>&nbsp;|&nbsp;"
shortNavBar=shortNavBar&"<b><a href='default.asp?"&QS_secCodeURL&"&amp;iId=" & encrypt(selectedPage.iId) &"&amp;item="& request("item") &"'>"&l("to")&"&nbsp;"&sName&"</a></b>"
shortNavBar=shortNavBar&"</div>"
end function
public function navbar(iTopicCount)
if not logon.authenticatedIntranet then
'alleen tonen als er nog geen posts zijn !
if iTopicCount=0 then
navbar=navbar& "<a href='default.asp?iId="&encrypt(selectedPage.iId)&"&amp;pageAction="&cloginIntranet&"&amp;item="& request("item") &"'><b>"&l("login")&"</b></a>"
end if
elseif allowPost then
navbar=navbar& "<a href='default.asp?"&QS_secCodeURL&"&amp;iId="&encrypt(selectedPage.iId)&"&amp;themeAction="&cPostTopic&"&amp;item="& request("item") &"'><b>"&l("newtopic")&"</b></a>"
end if
if allowSub then
'subscription to theme is allowed?
if iSubLevel=QS_theme_sublevel_theme then
if not isLeeg(navbar) then navbar=navbar& "&nbsp;|&nbsp;"
if logon.contact.bHasSubscribedToTheme(iId) then
navbar=navbar&"<a href='default.asp?"&QS_secCodeURL&"&amp;iId="&encrypt(selectedPage.iId)&"&amp;themeAction="&cUnSubscribeToTheme&"&amp;item="& Request("item") &"' onclick="& """" & "javascript:return confirm('"&l("areyousure")&"')" & """" &"><b>"&l("unsubscribefromtheme")&"</b></a>"
else
navbar=navbar&"<a href='default.asp?"&QS_secCodeURL&"&amp;iId="&encrypt(selectedPage.iId)&"&amp;themeAction="&cSubscribeToTheme&"&amp;item="& Request("item") &"' onclick="& """" & "javascript:return confirm('"&l("expl_subscribe_theme")&"')" & """" &"><b>"&l("subscribetotheme")&"</b></a>"
end if
end if
end if
'alleen navbar onder bepaalde voorwaarden...
if not (navbar="" and isLeeg(sPreviousPage) and isLeeg(sNextPage)) then
dim cnavbar, searchNav, subsNav
cnavbar=navbar
navbar="<div style='text-align:center;width:100%;font-size:0.9em;margin-bottom:10px;"
if selectedPage.iThemeID=iId then
navbar=navbar &"width:"&iWidth&";"
end if
if not isLeeg(cnavbar) then
searchNav="&nbsp;|&nbsp;"
end if
searchNav=searchNav & "<a href='default.asp?"&QS_secCodeURL&"&amp;iId="&encrypt(selectedPage.iId)&"&amp;themeAction=" & cSearchTheme & "'><b>" & l("search") & "</b></a>"
if QS_theme_sublevel_theme=iSubLevel then
if logon.contact.iId=iContactID and convertGetal(iContactID)<>0 then
searchNav=searchNav & "&nbsp;|&nbsp;<a href='default.asp?"&QS_secCodeURL&"&amp;iId="&encrypt(selectedPage.iId)&"&amp;themeAction="&clSubs&"'><b>" & l("subscriptions") & "</b></a>"
end if
end if
if not isLeeg(sNextPage) then
 'searchNav=searchNav& "&nbsp;|&nbsp;"
end if
navbar=navbar&"'>"&sPreviousPage & cnavbar & searchNav & sNextPage&"</div>"
end if
end function
public function posts
set posts=server.CreateObject ("scripting.dictionary")
if convertGetal(iId)=0 then exit function
dim rs, post, sql
sql="select iId from tblPost where (iPostID is null or iPostID=0) and iThemeID="& iId 
if iType=QS_theme_ts and convertGetal(iContactID)<>convertGetal(logon.contact.iId) then
sql=sql&" and iContactID="& logon.contact.iId
end if
if convertGetal(iContactID)<>convertGetal(logon.contact.iId) then
sql=sql&" and (bNeedsToBeValidated is null  or bNeedsToBeValidated="& getSQLBoolean(false) & ")"
end if
sql=sql & " order by dCreatedTS desc"
set rs=db.execute(sql)
totalNumber=0
dim continue
continue=true
'manual paging!
while not rs.eof and continue
if (startpage-1)*iPageSize<=totalNumber and totalNumber<iPageSize*startpage then
set post=new cls_post
post.pick(rs(0))
posts.Add post.iID,post
set post=nothing
end if
if totalNumber>=iPageSize*startpage	then continue=false
totalNumber=totalNumber+1
rs.movenext
wend
if not continue then
sNextPage="&nbsp;&nbsp;&nbsp;<a href='default.asp?iId="&encrypt(selectedpage.iId) & "&amp;startpage="&startpage+1&"&amp;item="& request("item") &"'>"&l("next")&"&nbsp;&gt;&gt;</a>"
end if
if startpage>1 then
sPreviousPage="<a href='default.asp?iId="&encrypt(selectedpage.iId) & "&amp;startpage="&startpage-1&"&amp;item="& request("item") &"'>&lt;&lt;&nbsp;"&l("previous")&"</a>&nbsp;&nbsp;&nbsp;"
end if
set rs=nothing
end function
public function copy()
if isNumeriek(iId) then
iId=null
sName=l("copyof") & " " & sName
sCode=""
save()
end if
end function
public function searchPosts(sValue)
set searchPosts=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
if not isLeeg(sValue) then
dim sql, rs, postObj, maxresults
maxresults=0
sql="select iId, iPostID from tblPost "
sql=sql & "where iThemeID=" & iId & " and (sBody like '%" & cleanup(left(sValue,100)) & "%' or sSubject like '%" & cleanup(left(sValue,100)) & "%') "
if iType=QS_theme_ts and convertGetal(iContactID)<>convertGetal(logon.contact.iId) then
sql=sql&" and iContactID="& logon.contact.iId
end if
sql=sql&" and (bNeedsToBeValidated is null  or bNeedsToBeValidated="& getSQLBoolean(false) & ") "
sql=sql & " order by dUpdatedTS desc"
set rs=db.execute(sql)
while not rs.eof and maxresults<200
set postObj=new cls_post
if convertGetal(rs(1))<>0 then
postObj.pick(rs(1))
else
postObj.pick(rs(0))
end if
if not searchPosts.Exists (postObj.iId) then
searchPosts.Add postObj.iId, postObj
maxresults=maxresults+1
end if
rs.movenext
set postObj=nothing
wend 
set rs=nothing
end if
end if
end function
private function buildSubs
dim cthemeSubs, contactSub
set cthemeSubs=themeSubs
buildSubs="<p><a href='default.asp?iId=" & encrypt(selectedPage.iId) & "'>" & l("back") & "</a></p><p><i>" & l("sublevel_theme") & ":</i></p><ul>"
for each contactSub in cthemeSubs
buildSubs=buildSubs&"<li>" & sanitize(cthemeSubs(contactSub).sNickName) & "&nbsp;-&nbsp;<a href='default.asp?iId=" & encrypt(selectedPage.iId) &"&amp;iThemeID=" & encrypt(iId) & "&amp;iContactID=" & encrypt(contactSub) & "&amp;themeAction=" & cdSub & "' onclick=""javascript:return confirm('"&l("areyousure")&"');"">" & l("delete") & "</a></li>"
next
set cthemeSubs=nothing
buildSubs=buildSubs&"</ul>"
buildSubs=buildSubs&"<p><i>" & l("sublevel_topic") & ":</i></p><ul>"
dim ctopicSubs, cts
set ctopicSubs=topicSubs
for each cts in ctopicSubs
buildSubs=buildSubs&"<li>" & ctopicSubs(cts) & "</li>"
next
buildSubs=buildSubs&"</ul>"
set ctopicSubs=nothing
end function
private function themeSubs
set themeSubs=server.CreateObject ("scripting.dictionary")
dim rs, subcontact
set rs=db.execute("SELECT tblThemeSubscription.iContactID FROM tblContact INNER JOIN tblThemeSubscription ON tblContact.iId = tblThemeSubscription.iContactID WHERE tblThemeSubscription.iThemeID=" & iId & " ORDER BY tblContact.sNickName asc")
while not rs.eof
set subcontact=new cls_contact
subcontact.pick(rs(0))
if not themeSubs.Exists (subcontact.iId) then themeSubs.Add subcontact.iId, subcontact
rs.movenext
set subcontact=nothing
wend 
set rs=nothing
end function
private function topicSubs
set topicSubs=server.CreateObject ("scripting.dictionary")
dim rs, subcontact, counter
set rs=db.execute("SELECT tblContact.sNickName, tblPost.sSubject FROM (tblContact INNER JOIN tblThemeTopicSubscription ON tblContact.iId = tblThemeTopicSubscription.iContactID) INNER JOIN tblPost ON (tblPost.iId = tblThemeTopicSubscription.iPostID) AND (tblContact.iId = tblPost.iContactID) WHERE tblPost.iThemeID="& iId & " order by tblContact.sNickName asc")
'                             0                     1                               2                3
counter=0
while not rs.eof
topicSubs.Add counter,sanitize(rs(0) & " - " & rs(1))
rs.movenext
counter=counter+1
wend 
set rs=nothing
end function
end class%>
