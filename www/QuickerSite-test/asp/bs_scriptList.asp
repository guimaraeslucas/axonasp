<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bHomeVBScript%><%if not customer.bApplication then Response.Redirect ("bs_constantList.asp")%><!-- #include file="bs_process.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Home)%><p align=center><%=getArtLink("bs_default.asp",l("pagelist"),"","","")%>&nbsp;&nbsp;
<%=getArtLink("bs_constantEdit.asp?iType=" &QS_VBScript,l("newitem"),"","","")%>&nbsp;&nbsp;
<%=getArtLink("bs_constantlist.asp",l("constants"),"","","")%></p><table align=center cellpadding="2"><tr><td><%dim cconstants, keyconstants
set cconstants=customer.constants
if cconstants.count>0 then%>ASP/VBScripts:<ul><%for each keyconstants in cconstants
if QS_VBScript=cconstants(keyconstants).iType then%><li><input onclick="javascript:this.select();" type="text" value="[<%=sanitize(cconstants(keyconstants).sConstant)%>]" size="20" />&nbsp;-&nbsp;<a href="bs_constantEdit.asp?iContentID=<%=encrypt(keyconstants)%>"><%=cconstants(keyconstants).sConstant%><%if not isLeeg(cconstants(keyconstants).sParameters) then Response.Write "("&cconstants(keyconstants).sParameters&")"%></a>&nbsp;-&nbsp;id: <%=keyconstants%>&nbsp;-&nbsp;<i><%=cconstants(keyconstants).statusString%></i></li><%end if
next%></ul><%set cconstants=nothing
end if%></td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
