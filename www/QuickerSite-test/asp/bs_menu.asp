
<%dim menu
set menu=new cls_menu
if isEmpty(convertStr(application(QS_CMS_cacheBOMenu & logon.currentPW))) or isleeg(application(convertStr(QS_CMS_cacheBOMenu & logon.currentPW))) then
application(QS_CMS_cacheBOMenu & logon.currentPW)=menu.getBOMenu(null,false)
end if
Response.Write application(QS_CMS_cacheBOMenu & logon.currentPW)
dim lossePaginas
set lossePaginas=menu.lossePaginas(false)
if lossePaginas.count>0 then
Response.Write l("freepages") & ":<ul>"
dim lp
for each lp in lossePaginas
Response.Write "<li><table cellpadding=1 cellspacing=0><tr><td valign=bottom>"
Response.Write lossePaginas(lp).getClickLink(true)	& "&nbsp;</td>"%><td><%=getIcon(l("freepage"),"free","#","javascript:return false;","free"&lp)%></td><%if not isLeeg(lossePaginas(lp).sApplication) and secondAdmin.bApplicationpath then%><td><%=getIcon(l("application"),"application","#","","application"&lp)%></td><%end if
if secondAdmin.bPagesAdd then%><td><%=getIcon(l("copyitem"),"copyItem","bs_default.asp?"&QS_secCodeURL&"&amp;btnaction=Copy&amp;iId="&EnCrypt(lp),"javascript:return copyItem();","copyItem"&lp)%></td><%end if
if secondAdmin.bPagesPW then
if isLeeg(lossePaginas(lp).sPw) then%><td><%=getIcon(l("managepw"),"unlock","bs_applyPw.asp?iId="&EnCrypt(lp),"","unlock"&lp)%></td><%else%><td><%=getIcon(l("managepw"),"lock","bs_applyPw.asp?iId="&EnCrypt(lp),"","lock"&lp)%></td><%end if
end if
if secondAdmin.bPagesMove then%><td><%=getIcon(l("moveitem"),"move","bs_selectPage.asp?"&QS_secCodeURL&"&amp;btnaction=Move&amp;iId="&EnCrypt(lp),"","move"&lp)%></td><%end if
Response.Write "</tr></table>"
next 
Response.Write "</ul>"
end if
set lossePaginas=nothing
set menu=nothing%>
