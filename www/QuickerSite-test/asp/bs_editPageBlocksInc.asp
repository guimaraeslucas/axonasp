
<%if isNumeriek(page.iID) and convertGetal(page.iListPageID)=0 and secondAdmin.bSetupPageElements then%><p align=center>-> <b><a class="bPopupFullWidthNoReload" href="bs_editPageBlocks.asp?iPageID=<%=encrypt(page.iId)%>">Edit Page Blocks</a></b> <-</p><%end if%>
