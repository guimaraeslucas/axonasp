
<%if secondAdmin.bPagePublish then
if not page.parentPage.bOnline then%><tr><td>&nbsp;</td><td><INPUT type="hidden" value="0" name=bOnline><%=l("cannotgooffline")%> /></td></tr><%elseIf not page.canGoOffline then%><tr><td>&nbsp;</td><td><INPUT type="hidden" value="1" name=bOnline /><%=l("cannotgoofflinehp")%></td></tr><%else %><tr><td class=QSlabel><%=l("online")%>:</td><td><input type=checkbox  style="BORDER:0px" name=bOnline value="1" 
<%if page.bOnline then 
response.write "checked "
if page.subPages(true).count>0 then%>onclick="javascript:if(!this.checked){alert('<%=l("warningoffline")%>')};"
<%end if
end if%> /></td></tr><%end if
end if%>
