<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bHomeConstants%><!-- #include file="bs_process.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Home)%><p align=center><%=getArtLink("bs_default.asp",l("pagelist"),"","","")%>&nbsp;&nbsp;
<%=getArtLink("bs_constantEdit.asp",l("newitem"),"","","")%><%if customer.bApplication and secondAdmin.bHomeVBScript then%>&nbsp;&nbsp;<%=getArtLink("bs_scriptlist.asp","ASP/VBScripts","","","")%><%end if%></p><table align=center cellpadding="2"><tr><td><%dim cconstants, keyconstants
set cconstants=customer.constants
dim formatList
set formatList=new cls_formatList
if cconstants.count>0 then%><%=l("constants")%>:<ul><%for each keyconstants in cconstants
if QS_VBScript<>cconstants(keyconstants).iType then%><li><input onclick="javascript:this.select()"  type="text" value="[<%=sanitize(cconstants(keyconstants).sConstant)%>]" size="20">&nbsp;-&nbsp;<a href="bs_constantEdit.asp?iContentID=<%=encrypt(keyconstants)%>"><%=cconstants(keyconstants).sConstant%></a>&nbsp;-&nbsp;<%=formatList.showSelected("single",cconstants(keyconstants).iType)%>&nbsp;-&nbsp;id: <%=keyconstants%>&nbsp;-&nbsp;<i><%=cconstants(keyconstants).statusString%></i></li><%end if
next%></ul><%end if
set formatList=nothing
set cconstants=nothing%></td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
