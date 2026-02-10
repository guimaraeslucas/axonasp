
	<%if secondAdmin.bPageRefresh then%><tr><td class=QSlabel><%=l("refresh")%>&nbsp;<%=l("after")%>:</td><td><select name=iReload><%=numberList(0,600,1,page.iReload)%></select>&nbsp;<i><%=l("seconds")%></i><br />(<%=l("redirectto")%>:&nbsp;<input type=text maxlength=255 size=40 name=sRedirectTo value="<%=quotRep(page.sRedirectTo)%>" />)</td></tr><%end if%>
