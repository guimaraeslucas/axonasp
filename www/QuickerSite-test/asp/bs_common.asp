
<%if isNumeriek(page.iParentID) then%><tr><td class=QSlabel><%=l("isattachedto")%>:</td><td><b><%=page.parentPage.sTitle%></b></td></tr><%end if%><tr><td valign=middle class=QSlabel><%=l("title")%>:*</td><td valign=middle><%if secondAdmin.bPageTitle then%><input required type=text maxlength=100 size=47 name=sTitle value="<%= quotRep(page.sTitle)%>" /><%else
response.write quotRep(page.sTitle)
end if%><input type=hidden name=bIntranet value="<%=page.bIntranet%>" /><!-- #include file="bs_preview.asp"--></td></tr><tr><td class=QSlabel><%=l("alternatetitle")%></td><td><%if secondAdmin.bPageTitle then%><input type=text maxlength=200 size=50 name=sPageTitle value="<%= quotRep(page.sPageTitle)%>" /><%else
response.write quotRep(page.sPageTitle)
end if%></td></tr>
