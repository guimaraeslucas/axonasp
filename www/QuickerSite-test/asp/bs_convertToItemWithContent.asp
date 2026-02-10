
<%if isNumeriek(page.iID) and not page.bHomepage  and secondAdmin.bPageBody then%><p align=center>-> <b><a href="bs_editItem.asp?iId=<%=encrypt(page.iId)%>"><%=l("convertToItemWithContent")%></a></b> <-</p><%end if%>
