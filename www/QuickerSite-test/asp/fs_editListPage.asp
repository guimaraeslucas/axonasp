<!-- #include file="begin.asp"-->


<%pagetoemail=true
editPage=true
dim EPSuccess, EPSaveSuccess
EPSuccess=false
EPSaveSuccess=false
dim accessToEP
accessToEP=false
dim getTperP,getBperP
set getTperP=logon.contact.getTper
set getBperP=logon.contact.getBper
logon.contact.createUserFilesFolder()
if convertBool(Request.Form ("postback")) then
checkCSRF()
if getTPerP.exists(selectedPage.iId) then
selectedPage.sTitle	= convertStr(Request.Form ("sTitle"))
end if 
if getBPerP.exists(selectedPage.iId) then
selectedPage.sValue	= convertStr(Request.Form ("sValue"))
selectedPage.bOpenOnload	= convertBool(Request.Form ("bOpenOnload"))
selectedPage.sOrderBY	= convertStr(Request.Form ("sOrderBY"))
selectedPage.bPushRSS	= convertBool(Request.Form ("bPushRSS"))
selectedPage.iLPOpenByDefault = convertGetal(Request.Form ("iLPOpenByDefault"))
end if
if getTPerP.exists(selectedPage.iId) or getBPerP.exists(selectedPage.iId) then
if selectedPage.Save then
message.Add("fb_saveOK")
EPSaveSuccess=true
end if
end if 
end if
dim orderByList
set orderByList=new cls_orderByList%><!-- #include file="includes/commonheader.asp"--><body style="color:#000;background-color:#FFF;background-image:none" onload="javascript:window.focus();"><%'if not EPSaveSuccess then%><form method="post" action="fs_editListPage.asp" name="mainform"><input type="hidden" name="iId" value="<%=encrypt(selectedPage.iId)%>" /><input type="hidden" name="postback" value="<%=true%>" /><%=QS_secCodeHidden%><table align="center"  style="background-color:#FFF"><%if getTPerP.exists(selectedPage.iId) then%><tr><td class="QSlabel"><%=l("title")%>:*</td><td><input type="text" size="40" maxlength="100" name="sTitle" value="<%=quotrep(selectedPage.sTitle)%>"></td></tr><%accessToEP=true
end if
if getBPerP.exists(selectedPage.iId) then%><tr><td class="QSlabel"><%=l("sortitemsby")%>:</td><td><select name=sOrderBY><%=orderByList.showSelected("option",selectedPage.sOrderBy)%></select></td></tr><tr><td class="QSlabel"><%=l("openitems")%></td><td><input type=checkbox  style="BORDER:0px" name=bOpenOnload value="1" <%=convertChecked(selectedPage.bOpenOnload)%> /></td></tr><%if not selectedPage.bOpenOnload then%><tr><td class="QSlabel"><%=l("ilpopenbydefault")%></td><td><select name="iLPOpenByDefault"><%=numberList(0,20,1,selectedPage.iLPOpenByDefault)%></select></td></tr><%end if%><tr><td class="QSlabel">Push RSS</td><td><input type=checkbox  style="BORDER:0px" name=bPushRSS value="1" <%=convertChecked(selectedPage.bPushRSS)%> /></td></tr><tr><td colspan="2"><%createFCKInstance selectedPage.sValue,"siteBuilder","sValue"%></td></tr><%accessToEP=true
end if%><tr><td class="QSlabel">&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class="QSlabel">&nbsp;</td><td><input type="submit" name="dummy" value="<% =l("save")%>" /></td></tr></table></form><%'end if%><%=message.showAlert()%><%=cPopup.dumpJS()%> 
</body></html><%if not accessToEP then Response.Redirect ("asp/noaccess.htm")%><%cleanUPASP%>
