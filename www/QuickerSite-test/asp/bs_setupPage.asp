<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bPagesAdd%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader("")%><!-- #include file="bs_process.asp"--><%dim fixedTypeList
set fixedTypeList=new cls_fixedTypeList
set fixedTypeList=fixedTypeList.list%><p align=center><%=l("selecttypeitem")%>:</p><form action="bs_setupPage.asp" method="post" name=mainform><INPUT type="hidden" value="<% =EnCrypt(page.iParentID) %>" name=iParentID /><INPUT type="hidden" value="continue" name=btnaction /><INPUT type="hidden" value="<%=request("bIntranet")%>" name=bIntranet /><table align="center" width="590" cellpadding="7" cellspacing="0"><%if isNumeriek(page.iParentID) then%><tr><td class=QSlabel style="width:140px"><%=l("isattachedto")%>:</td><td><b><%=page.parentPage.sTitle%></b></td></tr><tr><td class=QSlabel>&nbsp;</td><td><hr /></td></tr><%end if
dim fixedType,trunner
trunner=0
for each fixedType in fixedTypeList
if not (isNumeriek(page.iParentID) and fixedType=sb_lossePagina) then%><tr><td style="border-bottom:1px solid #EEEEEE" width=40 valign="top" align=right><input type=radio <%if trunner=0 and Request.Form("btnaction")<>"continue" then Response.Write "checked='checked' "%> name=fixedType value="<%=fixedType%>" /></td><td style="border-bottom:1px solid #EEEEEE"><b><%=fixedTypeList(fixedType)%></b><br /><%select case fixedType
case sb_item%><%=l("explsbitem")%><%case sb_container%><%=l("explsbcontainer")%><%case sb_externalUrl%><%=l("explsburl")%><%case sb_list%><%=l("explsblist")%><%case sb_lossePagina%><%=l("explsblp")%><%case sb_menugroup%><%=l("explsbmenugroup")%><%end select%></td></tr><%end if
trunner=trunner+1
next%><tr><td>&nbsp;</td><td><input class="art-button" type=submit name=dummy value="<%=l("continue")%>" /></td></tr></table></form><%if convertBool(request("bIntranet")) then%><!-- #include file="bs_backIntranet.asp"--><%else%><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
