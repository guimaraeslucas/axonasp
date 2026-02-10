
<%if customer.themes.count>0 and secondAdmin.bPageTheme then%><tr><td class="QSlabel"><%=l("theme")%>:</td><td><select name="iThemeID"><option value="">&nbsp;</option><%=customer.showSelectedTheme("option", page.iThemeID)%></select></td></tr><%end if%>
