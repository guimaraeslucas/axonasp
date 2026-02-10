<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bTheme%><!-- #include file="includes/header.asp"--><!-- #include file="includes/urlenCodeJS.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Intranet)%><%=getBOHeaderIntranet(btn_theme)%><%dim theme
set theme=new cls_theme
dim postBack
postBack=convertBool(Request.Form("postBack"))
if postBack then
theme.getRequestValues()
end if
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
theme.getRequestValues()
if theme.save then 
message.Add("fb_saveOK")
end if
case l("delete")
checkCSRF()
theme.remove
Response.Redirect ("bs_themesList.asp")
end select
dim themeTypeList
set themeTypeList=new cls_themeTypeList
dim contactSearch
set contactSearch=new cls_contactSearch
dim subLevelList
set subLevelList=new cls_theme_sublevelList
dim siteSearchTypeList
set siteSearchTypeList= new cls_siteSearchTypeList%><form action="bs_themeEdit.asp" method="post" name="mainform"><input type="hidden" name="iThemeId" value="<%=encrypt(theme.iID)%>" /><input type="hidden" value="<%=true%>" name="postback" /><%=QS_secCodeHidden%><table align="center" cellpadding="2"><tr><td colspan="2" class="header"><%=l("general")%>:</td></tr><tr><td class="QSlabel"><%=l("name")%>:*</td><td><input type="text" size="40" maxlength="255" name="sName" value="<%=quotRep(theme.sName)%>" /></td></tr><tr><td class="QSlabel"><%=l("code")%>:</td><td>[QS_THEME:<input type="text" size="10" maxlength="45" name="sCode" value="<%=quotRep(theme.sCode)%>" />]</td></tr><tr><td class="QSlabel"><%=l("type")%>:*</td><td><select onchange="javascript:document.mainform.submit();" name='iType'><%=themeTypeList.showSelected("option",theme.iType)%></select></td></tr><tr><td class="QSlabel"><%if theme.iType=QS_theme_cd or theme.iType=QS_theme_ts then%><%=l("moderator")%>:<%else%><%=l("blogger")%>:*<%end if%></td><td><select name="iContactID" <%if theme.iType=QS_theme_cd or theme.iType=QS_theme_ts then%> onchange="javascript:document.mainform.submit();" <%end if%>><option value=''><%=l("pleaseselect")%></option><%=contactSearch.showSelected(cs_write&","&cs_readwrite,theme.iContactID)%></select></td></tr><%if (theme.iType=QS_theme_cd or theme.iType=QS_theme_ts) and convertGetal(theme.iContactID)<>0 then%><tr><td class="QSlabel"><%=l("fowardpoststomoderator")%></td><td><input type="checkbox" name="bForwardPostsToModerator" value="1" <%=convertChecked(theme.bForwardPostsToModerator)%> /></td></tr><%end if
if convertGetal(theme.iContactID)<>0 then%><tr><td class="QSlabel"><%=l("needsvalidation")%></td><td><input type="checkbox" name="bValidation" value="1" <%=convertChecked(theme.bValidation)%> /></td></tr><%end if%><%dim iSubLevelHF
if theme.iType=QS_theme_ts then
iSubLevelHF="<input type=""hidden"" name=""iSubLevel"" value=""" & QS_theme_sublevel_authortopic & """ />"
else%><tr><td class="QSlabel"><%=l("higestlevelofsub")%></td><td><select name="iSubLevel"><%=subLevelList.showSelected("option",theme.iSubLevel)%></select></td></tr><%end if%><tr><td class="QSlabel"><%=l("includetopics/repliesinsite-search?")%></td><td><select name="iSearchType"><%=siteSearchTypeList.showSelected("option",theme.iSearchType)%></select></td></tr><tr><td class="QSlabel"><%=l("online")%></td><td><input type="checkbox" name="bOnline" value="1" <%=convertChecked(theme.bOnline)%> /></td></tr><tr><td class="QSlabel"><%=l("locktheme")%></td><td><input type="checkbox" name="bLocked" value="1" <%=convertChecked(theme.bLocked)%> /></td></tr><tr><td class="QSlabel"><%=l("allowhtml")%></td><td><input onclick="javascript:document.mainform.submit();" type="checkbox" name="bAllowHTML" value="1" <%=convertChecked(theme.bAllowHTML)%> /></td></tr><%if theme.bAllowHTML then%><tr><td class="QSlabel"><%=l("allowupload")%></td><td><input type="checkbox" name="bUpload" value="1" <%=convertChecked(theme.bUpload)%> /></td></tr><tr><td class="QSlabel"><%=l("allowembeddingobjects?")%></td><td><input type="checkbox" name="bEmbed" value="1" <%=convertChecked(theme.bEmbed)%> /></td></tr><%end if%><tr><td class="QSlabel">Smilies?</td><td><input type="checkbox" name="bSmileys" value="1" <%=convertChecked(theme.bSmileys)%> /></td></tr><tr><td class="QSlabel">Enable attachments in topics/replies?</td><td><input type="checkbox" name="bFileUploads" value="1" <%=convertChecked(theme.bFileUploads)%> /> (will always be DISABLED for anonymous users!)</td></tr><tr><td class="QSlabel"><%=l("showtopictitlesonly")%></td><td><input type="checkbox" name="bCompactList" value="1" <%=convertChecked(theme.bCompactList)%> /></td></tr><%if not theme.bAllowHTML then%><tr><td class="QSlabel"><%=l("limittopicsto")%></td><td><select name="iLimitTopicTo"><%=numberList(500,9000,100,theme.iLimitTopicTo)%></select> <i><%=l("characters")%></i></td></tr><%end if%><tr><td class=QSlabel><%=l("allowreplies")%></td><td><input type=checkbox onclick="javascript:document.mainform.submit();" name="bAllowComments" value="1" <%=convertChecked(theme.bAllowComments)%> /></td></tr><%if theme.bAllowComments then%><tr><td class="QSlabel">Allow anonymous replies?</td><td><input onclick="javascript:document.mainform.submit();" type="checkbox" name="bAllowAP" value="1" <%=convertChecked(theme.bAllowAP)%> /></td></tr><%if convertBool(theme.bAllowAP) then%><tr><td class="QSlabel">Label 'your name':*</td><td><input type="text" size="40" maxlength="50" name="sLabelYourName" value="<%=quotRep(theme.sLabelYourName)%>" /></td></tr><%end if%><tr><td class="QSlabel"><%=l("limitrepliesto")%></td><td><select name=iLimitReplyTo><%=numberList(150,3500,50,theme.iLimitReplyTo)%></select> <i><%=l("characters")%></i></td></tr><%end if%><tr><td class=QSlabel><%=l("topicsperpage")%>:</td><td><select name="iPageSize"><%=numberList(1,100,1,theme.iPageSize)%></select></td></tr><tr><td class=QSlabel><%=l("coloreventopics")%>:</td><td><input type="text" id="sColorEven" name="sColorEven" value="<%=quotrep(theme.sColorEven)%>" /><%=JQColorPicker("sColorEven")%></td></tr><tr><td class=QSlabel><%=l("coloruneventopics")%>:</td><td><input type="text" id="sColorUnEven" name="sColorUnEven" value="<%=quotrep(theme.sColorUnEven)%>" /><%=JQColorPicker("sColorUnEven")%></td></tr><tr><td class=QSlabel><%=l("width")%>:</td><td><select name=iWidth><%=numberList(150,1500,1,theme.iWidth)%></select> <i>px - <%=l("advised")%>: <%=theme.recomWidth%> px</i></td></tr><%dim bPushRSSHF
if theme.iType=QS_theme_ts then
bPushRSSHF="<input type=""hidden"" name=""bPushRSS"" value=""0"" />"
else%><tr><td class=QSlabel>Push RSS?</td><td><input type=checkbox name="bPushRSS" value="1" <%=convertChecked(theme.bPushRSS)%> /></td></tr><%end if%><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /><%if isNumeriek(theme.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><br /><br /><%end if%></td></tr><tr><td colspan=2><hr /></td></tr><tr><td class=QSlabel><%=l("notificationtopicsubject")%>*</td><td><input type=text size=40 maxlength=255 name="sTopicSubjectNotification" value="<%=quotRep(theme.sTopicSubjectNotification)%>" /></td></tr><tr><td class=QSlabel width=250><%=l("notificationtopicbody")%>*</td><td><%createFCKInstance theme.sTopicBodyNotification,"siteBuilderRichTextSmall","sTopicBodyNotification"%></td></tr><tr><td class=QSlabel><%=l("notificationsubject")%>*</td><td><input type=text size=40 maxlength=255 name="sSubjectNotification" value="<%=quotRep(theme.sSubjectNotification)%>" /></td></tr><tr><td class=QSlabel width=250><%=l("notificationbody")%>*</td><td><%createFCKInstance theme.sBodyNotification,"siteBuilderRichTextSmall","sBodyNotification"%></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr></table><%=bPushRSSHF%><%=iSubLevelHF%></form><%Response.Flush 
if convertGetal(theme.iId)<>0 then
dim fsearch
set fsearch=new cls_fullSearch
fsearch.pattern="\[QS_THEME:+("&theme.sCode&")+[\]]"
dim result
result=fsearch.search()
if not isLeeg(result) then
Response.Write "<p align=center>"&l("wherethemeused")&"</p>"
Response.Write result
end if
set fsearch=nothing
end if%><table align=center><tr><td align=center>-> <b><a href="bs_themesList.asp"><%=l("back")%></a></b> <-</td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
