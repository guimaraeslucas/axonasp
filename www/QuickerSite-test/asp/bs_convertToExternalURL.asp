
<%if isNumeriek(page.iID) and not page.bHomepage  and secondAdmin.bPageBody then%><p align=center>-> <b><a href="bs_editExternalURL.asp?iId=<%=encrypt(page.iId)%>"><%=l("convertToExternalURL")%></a></b> <-</p><%end if%>
