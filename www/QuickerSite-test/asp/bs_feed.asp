
<%if customer.feeds.count>0 and secondAdmin.bPageFeed and convertGetal(page.ifeedId)<>0 then%>
<tr><td class=QSlabel><%=l("feed")%>:</td>
<td><select name=ifeedId><option value="">&nbsp;</option>
<%=customer.showSelectedfeed("option", page.ifeedId)%>
</select></td></tr>
<%end if%>
