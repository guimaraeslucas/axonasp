<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bIntranetContacts%><%dim contact
set contact=new cls_contact
dim postback
postback=convertBool(Request.Form ("postback"))
if postback then
contact.savePermissions Request.Form ("iPageEditID"),Request.Form ("iTitleEditID"),Request.Form ("bLPID")
end if
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
if contact.savePermissions(Request.Form ("bBodyID"),Request.Form ("bTitleID"),Request.Form ("bLPID")) then
message.Add("fb_saveOK")
end if
end select
dim getPRMenu
set getPRMenu=new cls_menu
dim lossePaginas,lossePaginasI
set lossePaginas=getPRMenu.lossePaginas(false)
set lossePaginasI=getPRMenu.lossePaginas(true)
dim getTPer
set getTPer=contact.getTPer
dim getBPer
set getBPer=contact.getBPer
dim getLPer
set getLPer=contact.getLPer%><!-- #include file="includes/commonheader.asp"--><body style="background-color:#FFF"><form action="bs_contactPage.asp" method="post" name=mainform><%=QS_secCodeHidden%><input type="hidden" name="iContactID" value="<%=encrypt(contact.iID)%>" /><input type="hidden" value="<%=true%>" name="postback" /><table align="center" cellpadding="2"><tr><td colspan=2 class=header><%=l("managepermissions")%>:</td></tr><tr><td class=QSlabel><%=l("nickname")%>:</td><td><%=quotRep(contact.sNickname)%></td></tr><tr><td class=QSlabel><%=l("email")%>:</td><td><%=quotRep(contact.sEmail)%></td></tr><tr><td colspan=2><hr /></td></tr><tr><td class=QSlabel valign="top"><%=l("managepermissions")%>:</td><td><%=getPRMenu.getPRMenu(null,false,getTPer,getBPer,getLPer)%><br /><strong><%=l("freepages")%>:</strong><%dim lp
if lossePaginas.count>0 then%><ul><%for each lp in lossePaginas %><li><b><%=lossePaginas(lp).sTitle%>: </b> <input <%=convertChecked(getTPer.exists(lp))%> type="checkbox" value="iPageIDTitle<%=lp%>" name="bTitleID"> <%=l("edittitle")%> <input <%=convertChecked(getBPer.exists(lp))%>  type="checkbox" value="iPageIDBody<%=lp%>" name="bBodyID" /><%=l("editpage")%> </li><%next%></ul><%end if%></td></tr><tr><td class=QSlabel valign="top"><%=l("managepermissions") & " " & l("intranet") %> :</td><td><%=getPRMenu.getPRMenu(null,true,getTPer,getBPer,getLPer)%><br /><strong><%=l("freepages") & " " & l("intranet")%>:</strong><%if lossePaginasI.count>0 then%><ul><%for each lp in lossePaginasI %><li><b><%=lossePaginasI(lp).sTitle%>: </b> <input <%=convertChecked(getTPer.exists(lp))%> type="checkbox" value="iPageIDTitle<%=lp%>" name="bTitleID" /> <%=l("edittitle")%> <input  <%=convertChecked(getBPer.exists(lp))%>  type="checkbox" value="iPageIDBody<%=lp%>" name="bBodyID" /><%=l("editpage")%> </li><%next%></ul><%end if%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /></td></tr></table></form><%=message.showAlert()%><%=cPopup.dumpJS()%> 
</body></html><%cleanUPASP%>
