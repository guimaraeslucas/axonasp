
<%if isNumeriek(page.iParentID) then%><tr><td class=QSlabel><%=l("isattachedto")%>:</td><td><b><%=page.parentPage.sTitle%></b></td></tr><%end if%><tr><td class=QSlabel><%=l("title")%>:*</td><td><%if secondAdmin.bPageTitle then%><input type=text maxlength=100 size=40 name=sTitle value="<%= quotRep(page.sTitle)%>" /><%else
response.write quotRep(page.sTitle)
end if%><input type=hidden name=bIntranet value="<%=page.bIntranet%>" /></td></tr>
