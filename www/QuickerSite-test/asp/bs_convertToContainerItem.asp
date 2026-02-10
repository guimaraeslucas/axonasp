
<%if isNumeriek(page.iID) and not page.bHomepage  and secondAdmin.bPageBody then%><p align=center>-> <b><a href="bs_editContainer.asp?iId=<%=encrypt(page.iId)%>"><%=l("convertToContainerItem")%></a></b> <-</p><%end if%>
