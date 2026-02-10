<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bHomeConstants%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader("")%><%dim constant
set constant=new cls_constant
dim postback
postback=convertBool(Request.Form ("postback"))
if postback then constant.getRequestValues()
select case Request ("btnaction")
case l("save")
checkCSRF()
constant.getRequestValues()
if constant.save then 
message.Add("fb_saveOK")
end if
case l("delete")
checkCSRF()
constant.remove
if constant.iType=QS_VBscript then
Response.Redirect ("bs_scriptlist.asp")
else
Response.Redirect ("bs_constantlist.asp")
end if
end select
dim formatList
set formatList=new cls_formatList%><p align="center"><%if constant.iType=QS_VBSCRIPT then 
Response.Write getArtLink("bs_scriptlist.asp",l("back"),"","","")
else 
Response.Write getArtLink("bs_constantlist.asp",l("back"),"","","")
end if%></p><form action="bs_constantEdit.asp" method="post" name=mainform><%=QS_secCodeHidden%><input type="hidden" name=iContentId value="<%=encrypt(constant.iID)%>" /><input type="hidden" value="<%=true%>" name="postback" /><input type="hidden" value="<%=l("save")%>" name="btnaction" /><table align="center" width="500" cellpadding="2"><tr><td colspan=2 class=header><%=l("general")%>:</td></tr><tr><td class=QSlabel><%=l("constant")%>:*</td><td><input type=text size=40 maxlength=20 name="sConstant" value="<%=quotRep(constant.sConstant)%>" /><%if not isLeeg(constant.iId) then%>&nbsp;<a href="#" onclick="javascript:if (confirm('<%=l("areyousure")%>')){location.assign('bs_constantEdit.asp?<%=QS_secCodeURL%>&amp;iContentId=<%=encrypt(constant.iId)%>&amp;btnaction=<%=l("delete")%>')};"><img alt='<%=l("delete")%>' src="<%=C_DIRECTORY_QUICKERSITE%>/fixedImages/dustbin.gif" border=0 align=top /></a><%end if%></td></tr><tr><td class=QSlabel><%=l("type")%>:</td><td><select onchange="javascript:document.mainform.btnaction.value='';document.mainform.submit();" name=iType><%=formatList.showSelected("option",constant.iType)%></select></td></tr><tr><td class=QSlabel><%=l("online")%>&nbsp;<%=l("from")%>:</td><td><input type="text" id="dOnlinefrom" name="dOnlinefrom" value="<%=convertEuroDate(constant.dOnlineFrom)%>" /><%=JQDatePicker("dOnlinefrom")%><i><%=l("untill")%>:</i> 
<input type="text" id="dOnlineUntill" name="dOnlineUntill" value="<%=convertEuroDate(constant.dOnlineUntill)%>" /><%=JQDatePicker("dOnlineUntill")%></td></tr><%select case constant.iType 
case QS_textonly%><tr><td class=QSlabel><%=l("textonly")%>:*</td><td><textarea cols=110 rows=20 name=sValue><%=quotrep(constant.sValue)%></textarea></td></tr><%case QS_html%><tr><td align=center colspan=2><%createFCKInstance constant.sValue,"siteBuilder","sValue"%></td></tr><%case QS_VBScript%><tr><td class=QSlabel valign=top>ASP/VBScript:*</td><td><p>You can develop an ASP/VBScript function here. Anything you <b>Response.Write</b> (or directly assign to <b>CustomFunction</b>), will be included exactly where you add this constant. Make sure you <b>test</b> the function before saving it! If a runtime error occurs, it will be only be displayed in the test-mode, never in your site!</p><a name="sf"><b>function CustomFunction(Parameters: <input type=text size=50 maxlength=200 value="<%=sanitize(constant.sParameters)%>" name=sParameters />)</b></a><br />&nbsp;&nbsp;<textarea style="word-wrap:normal; overflow-x:auto; overflow-y:auto" cols="110" rows="20" name="sValue"><%=quotrep(constant.sValue)%></textarea><br /><a name="ef"><b>end function</b></a></td></tr><tr><td class=QSlabel valign=top>Global Code (VBScript Functions and Classes):</td><td><textarea style="word-wrap:normal; overflow-x:auto; overflow-y:auto" cols="105" rows="20" name="sGlobal"><%=quotrep(constant.sGlobal)%></textarea></td></tr><%end select%><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><%if constant.iType<>QS_html then%><tr><td class=QSlabel>&nbsp;</td><td><%if QS_VBScript=constant.iType then%><input class="art-button" type=button onclick="javascript:document.scriptform.customScript.value=document.mainform.sValue.value;document.scriptform.sGlobal.value=document.mainform.sGlobal.value;document.scriptform.sParameters.value=document.mainform.sParameters.value;document.scriptform.submit();" value="Test!" name=btnTest />&nbsp;<%end if%><input <%if QS_VBScript=constant.iType then%> onclick="javascript: if (confirm('Are you sure? Did you TEST the function?')){return true;}else{return false;};"<%end if%> class="art-button" type=submit value="<%=l("save")%>" name=dummyAction /></td></tr><%else%><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit value="<%=l("save")%>" name=dummyAction /></td></tr><%end if %></table></form><%Response.Flush 
if convertGetal(constant.iId)<>0 then
dim fsearch
set fsearch=new cls_fullSearch
fsearch.pattern="\[+("&constant.sConstant&")+[\(]+[\S| ]+[\)]+[\]]|\[+("&constant.sConstant&")+[\]]"
dim result
result=fsearch.search()
if not isLeeg(result) then
Response.Write "<p align=center>"&l("whereused")&"</p>"
Response.Write result
end if
set fsearch=nothing
end if
if QS_VBScript=constant.iType then%><form target="<%=generatePassword()%>" action="bs_constantTest.asp" method="post" name="scriptform"><input type="hidden" name="customScript" /><input type="hidden" name="sParameters" /><input type="hidden" name="sGlobal" /><%=QS_secCodeHidden%></form><%end if
set constant=nothing%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
