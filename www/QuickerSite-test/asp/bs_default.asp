<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><!-- #include file="bs_process.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Home)%><p align=center><%=getArtLink("bs_default.asp",l("pagelist"),"","","")%><%if secondAdmin.bPagesAdd then%>&nbsp;&nbsp;<%=getArtLink("bs_setupPage.asp",l("newpage"),"","","")%><%end if%><%if secondAdmin.bHomeConstants then%>&nbsp;&nbsp;<%=getArtLink("bs_constantlist.asp",l("constants"),"","","")%><%end if%><%if customer.bApplication and secondAdmin.bHomeVBScript then%>&nbsp;&nbsp;<%=getArtLink("bs_scriptlist.asp","ASP/VBScripts","","","")%><%end if%><%if logon.currentPW<>customer.secondAdmin.sPassword then%>&nbsp;&nbsp;<%=getArtLink("bs_search.asp",l("search"),"","","")%><%end if%><%if customer.pagesTobeValidated.count>0 then%>&nbsp;&nbsp;<%=getArtLink("bs_validatepages.asp","Validations","","","")%><%end if%></p><table align=center><tr><td><!-- #include file="bs_menu.asp"--></td></tr></table><script type="text/javascript"><!--
function copyItem(){
 
 if (confirm('<%=l("areyousuretocopy")%>')){
return true;
 }
 else
 {
 return false;
 }
}
//--></script><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
