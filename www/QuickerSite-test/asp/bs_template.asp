
<%if customer.templates.count>0 and secondAdmin.bPageTemplate then%><tr><td class="QSlabel"><%=l("template")%>:</td><td><select name="iTemplateId"><option value="">&nbsp;</option><%=customer.showSelectedtemplate("option", page.iTemplateId)%></select></td></tr><%end if%>
