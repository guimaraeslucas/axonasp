<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bShoppingCart%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_qsc)%><!-- #include file="bs_shoppingcartmenu.asp"--><%dim shopMake
set shopMake=new cls_shopMake
shopMake.pick(decrypt(request("iShopMakeID")))
dim postback
postback=convertBool(request.form("postback"))
if postback then
select case request.form("btnaction")
case l("save")
shopMake.sName=request.form("sName")
shopMake.bOnline=request.form("bOnline")
shopMake.sLogo=request.form("sLogo")
if shopMake.save then response.redirect ("bs_shoppingcart.asp")
case l("delete")
shopMake.delete
response.redirect ("bs_shoppingcart.asp")
end select
end if%><form action="bs_shopMake.asp" method="post" name="mainform"><%=QS_secCodeHidden%><input type="hidden" name="iShopMakeID" value="<%=encrypt(shopMake.iID)%>" /><input type="hidden" value="<%=true%>" name="postback" /><table align=center cellpadding="4"><tr><td class="QSlabel">Make name:*</td><td><input type=text size=40 maxlength=250 name="sName" value="<%=sanitize(shopMake.sName)%>" /></td></tr><tr><td class="QSlabel">Url to Make logo:*</td><td><input type=text size=40 maxlength=250 name="sLogo" value="<%=sanitize(shopMake.sLogo)%>" /></td></tr><tr><td class=QSlabel><%=l("online")%></td><td><input type=checkbox name="bOnline" value="1" <%=convertChecked(shopMake.bOnline)%> /></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /><%if isNumeriek(shopMake.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><br /><br /><%end if%></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
