
<%if customer.bCatalog and customer.catalogs.count>0 and secondAdmin.bPageCatalog then%><tr><td class=QSlabel><%=l("catalog")%>:</td><td><select name=iCatalogID><option value="">&nbsp;</option><%=customer.showSelectedCatalog("option", page.iCatalogID)%></select></td></tr><%end if%>
