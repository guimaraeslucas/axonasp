
<%if isNumeriek(page.iID) and secondAdmin.bPagesAdd and not page.bLossePagina then%><p align=center>-> <b><a href="bs_setupPage.asp?bIntranet=<%=convertBool(page.bIntranet)%>&amp;iParentid=<%=encrypt(page.iId)%>"><%=l("addnewitem")%></a></b> <-</p><%end if%>
