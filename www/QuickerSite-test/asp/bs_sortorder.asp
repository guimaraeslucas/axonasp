
<%if not page.bLossePagina then
if secondAdmin.bPagesMove and secondAdmin.bPageOrder then
dim pageRang, allRang
if convertGetal(page.iId)=0 then 
pageRang=page.siblings+1
allRang=page.siblings+1
else
pageRang=page.iRang
allRang=page.siblings
end if%><tr><td class=QSlabel><%=l("placeinmenu")%>:*</td><td><select name="iRang"><%=numberList(1,allRang,1,pageRang)%></select></td></tr><%end if%><%end if%>
