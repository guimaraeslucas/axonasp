
<%dim alignList
set alignList=new cls_alignList
if customer.forms.count>0 and secondAdmin.bPageForm then%><tr><td class=QSlabel><%=l("form")%>/<%=l("align")%>:</td><td><select name="iFormID"><option value="">&nbsp;</option><%=customer.showSelectedForm("option", page.iFormID)%></select>&nbsp;<select name=sFormAlign><%=alignList.showSelected("option", page.sFormAlign)%></select></td></tr><%end if%>
