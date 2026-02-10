<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bPagesMove%><!-- #include file="bs_process.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader("")%><%dim subPages
set subPages=page.subPages(false)%><table align=center cellpadding="2"><tr><td height=25><%=l("click")%></td><td><%=getIcon(l("insertItem"),"insertItem"&tdir,"#","","insertMain")%></td><td><%=l("toinsertitem")%>: '<b><%=page.sTitle%></b>'</td></tr></table><%if isNumeriek(page.iParentID) or convertBool(page.bLossePagina) then%><p align=center><a href="bs_selectPage.asp?<%=QS_secCodeURL%>&amp;btnaction=Insert&amp;iId=<%=enCrypt(page.iId)%>" title='<%=l("newmainitem")%>'><%=l("makemainitemof")%> '<b><%=page.sTitle%></b>'</a></p><%end if%><br /><table align=center><tr><td><%dim menu
set menu=new cls_menu
menu.bOnline=false
Response.Write menu.getReplaceMenu(null,subPages,page)
set menu=nothing%></td></tr></table><br /><%if convertBool(page.bIntranet) then%><!-- #include file="bs_backIntranet.asp"--><%else%><%end if%><!-- #include file="bs_endBack.asp"--><script type="text/javascript"><!--
function insertItem(){
 
 if (confirm('<%=l("suretomoveitem")%>')){
return true;
 }
 else
 {
 return false;
 }
}
//--></script><!-- #include file="includes/footer.asp"-->
