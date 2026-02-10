<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bFeed%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_feed)%><%dim feed
set feed=new cls_feed
dim postBack
postBack=convertBool(Request.Form("postBack"))
if postBack then feed.getRequestValues()
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
feed.getRequestValues()
if feed.save then 
Response.Redirect ("bs_feedList.asp")
end if
case l("delete")
checkCSRF()
feed.remove
Response.Redirect ("bs_feedList.asp")
end select
dim urlTypeShortList
set urlTypeShortList=new cls_urlTypeShortList%><!-- #include file="bs_feedBack.asp"--><form action="bs_feedEdit.asp" method="post" name=mainform><input type=hidden name=ifeedId value="<%=encrypt(feed.iID)%>" /><input type="hidden" value="<%=true%>" name=postback /><%=QS_secCodeHidden%><table align=center cellpadding="2"><tr><td colspan=2 class=header><%=l("general")%>:</td></tr><tr><td class=QSlabel><%=l("name")%>:*</td><td><input type=text size=40 maxlength=255 name="sName" value="<%=quotRep(feed.sName)%>" /></td></tr><tr><td class=QSlabel><%=l("code")%>:</td><td>[QS_FEED:<input type=text size=10 maxlength=45 name="sCode" value="<%=quotRep(feed.sCode)%>" />]</td></tr><tr><td class=QSlabel><%=l("url")%>:*</td><td><select name="sPrefixUrl"><%=urlTypeShortList.showSelected("option",feed.sPrefixUrl)%></select>&nbsp;<input type=text size=66 maxlength=255 name="sUrl" value="<%=quotRep(feed.sUrl)%>" /></td></tr><tr><td class=QSlabel>Urls:<sup>(1)</sup></td><td><textarea name="sUrls" style="word-wrap:normal; overflow-x:auto; overflow-y:auto" cols="80" rows="5"><%=quotRep(feed.sUrls)%></textarea><br /><small><sup>(1)</sup> <i>only use in case you need to combine multiple feeds</i></small></td></tr><tr><td class=QSlabel><%=l("maxitems")%>:</td><td><select name="iMaxItems"><%=numberList(1,100,1,feed.iMaxItems)%></select></td></tr><tr><td class=QSlabel><%=l("showitemsatrandom")%></td><td><input type=checkbox name="bRandom" value="1" <%=convertChecked(feed.bRandom)%> /></td></tr><tr><td class=QSlabel><%=l("enablejs")%></td><td><input type=checkbox name="bEnableJS" value="1" <%=convertChecked(feed.bEnableJS)%> /></td></tr><tr><td class=QSlabel><%=l("cache")%>:</td><td><select name="iCache"><%=numberList(0,14400,30,feed.iCache)%></select>&nbsp;<%=l("seconds")%></td></tr><tr><td colspan=2 class=header><%=l("advanced")%>:</td></tr><tr><td class=QSlabel><%=l("usecustomtemplate")%></td><td><input type=checkbox name="bTemplate" onclick="javascript:document.mainform.submit();" value="1" <%=convertChecked(feed.bTemplate)%> /></td></tr><%if not convertBool(feed.bTemplate) then%><tr><td class=QSlabel><%=l("showitemtitle")%></td><td><input type=checkbox name="bShowTitle" onclick="javascript:document.mainform.submit();" value="1" <%=convertChecked(feed.bShowTitle)%> /></td></tr><%if feed.bShowTitle then%><tr><td class=QSlabel><%=l("linkonitemtitle")%></td><td><input type=checkbox name="bLinkOnTitle" value="1" <%=convertChecked(feed.bLinkOnTitle)%> /></td></tr><tr><td class=QSlabel><%=l("openlinkinnw")%></td><td><input type=checkbox name="bOpenLinkInNW" value="1" <%=convertChecked(feed.bOpenLinkInNW)%> /></td></tr><%end if%><tr><td class=QSlabel><%=l("showitemdate")%></td><td><input type=checkbox name="bShowDate" value="1" <%=convertChecked(feed.bShowDate)%> /></td></tr><tr><td class=QSlabel><%=l("showauthorinformation")%></td><td><input type=checkbox name="bShowAuthor" value="1" <%=convertChecked(feed.bShowAuthor)%> /></td></tr><tr><td class=QSlabel><%=l("showitemcategory")%></td><td><input type=checkbox name="bShowCategory" value="1" <%=convertChecked(feed.bShowCategory)%> /></td></tr><%else%><tr><td class=QSlabel valign=top><%=l("htmlbefore")%>:</td><td><input type=text size=55 maxlength=254 name="sHTMLBefore" value="<%=quotRep(feed.sHTMLBefore)%>" /></td></tr><tr><td class=QSlabel valign=top><%=l("template")%>:<br /><input type=text size=16 value="{AUTHOR}" onclick="javascript:this.select();" /><br /><input type=text size=16 value="{CATEGORY}" onclick="javascript:this.select();" /><br /><input type=text size=16 value="{COUNTER}" onclick="javascript:this.select();" /><br /><input type=text size=16 value="{DATE}" onclick="javascript:this.select();" /><br /><input type=text size=16 value="{DESCRIPTION}" onclick="javascript:this.select();" /><br /><input type=text size=16 value="{IMAGE}" onclick="javascript:this.select();" /><br /><input type=text size=16 value="{ENCLOSURE}" onclick="javascript:this.select();" /><br /><input type=text size=16 value="{LINK}" onclick="javascript:this.select();" /><br /><input type=text size=16 value="{TITLE}" onclick="javascript:this.select();" /></td><td><textarea cols=90 rows=14 name="sTemplate"><%=quotRep(feed.sTemplate)%></textarea></td></tr><tr><td class=QSlabel valign=top><%=l("htmlafter")%>:</td><td><input type=text size=55 maxlength=254 name="sHTMLAfter" value="<%=quotRep(feed.sHTMLAfter)%>" /></td></tr><%end if%>
<tr><td class=QSlabel>Limit title to:</td><td><select name="iTitleLimitTo"><%=numberList(0,100,1,feed.iTitleLimitTo)%></select>&nbsp;<%=l("characters")%> (0: <%=l("nolimit")%>)</td></tr>
<tr><td class=QSlabel><%=l("limitto")%>:</td><td><select name="iLimitTo"><option value="<%=QS_feedNoText%>"><%=l("notext")%></option><%=numberList(0,2500,5,feed.iLimitTo)%></select>&nbsp;<%=l("characters")%> (0: <%=l("nolimit")%>)</td></tr>
<tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name="btnaction" value="<% =l("save")%>" /><%if isNumeriek(feed.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><br /><br /><%end if%></td></tr></table></form><%Response.Flush 
if convertGetal(feed.iId)<>0 then
dim fsearch
set fsearch=new cls_fullSearch
fsearch.pattern="\[QS_FEED:+("&feed.sCode&")+[\]]"
dim result
result=fsearch.search()
if not isLeeg(result) then
Response.Write "<p align=center>"&l("wherefeedused")&"</p>"
Response.Write result
end if
set fsearch=nothing
end if%><!-- #include file="bs_feedBack.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
