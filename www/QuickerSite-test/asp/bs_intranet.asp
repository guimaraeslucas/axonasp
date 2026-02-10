<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bIntranet%><!-- #include file="bs_process.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Intranet)%><%=getBOHeaderIntranet(btn_IntranetI)%><p align=center><a href="bs_setupPage.asp?bIntranet=<%=true%>"><b><%=l("newitem")%></b></a></p><table align=center><tr><td><!-- #include file="bs_menuIntranet.asp"--></td></tr></table><script type="text/javascript"><!--
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
