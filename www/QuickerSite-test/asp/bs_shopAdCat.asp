<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bShoppingCart%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_qsc)%><!-- #include file="bs_shoppingcartmenu.asp"--><%dim shopCat
set shopCat=new cls_shopCategory
shopCat.pick(decrypt(request("iShopCatID")))
dim postback
postback=convertBool(request.form("postback"))
if postback then
select case request.form("btnaction")
case l("save")
shopCat.sName=request.form("sName")
shopCat.bOnline=request.form("bOnline")
shopCat.iParentCatID=request.form("iParentCatID")
if shopCat.save then response.redirect ("bs_shoppingcart.asp")
case l("delete")
shopCat.delete
response.redirect ("bs_shoppingcart.asp")
end select
end if%><form action="bs_shopAdCat.asp" method="post" name="mainform"><%=QS_secCodeHidden%><input type="hidden" name="iShopCatID" value="<%=encrypt(shopCat.iID)%>" /><input type="hidden" value="<%=true%>" name="postback" /><table align=center cellpadding="4"><tr><td class="QSlabel">Category name:*</td><td><input type=text size=40 maxlength=250 name="sName" value="<%=sanitize(shopCat.sName)%>" /></td></tr><%if shopCat.bShowParentCatDropDown then%><tr><td class="QSlabel">Parent Category:</td><td><select name=iParentCatID><option value="0">&nbsp;</option><%=shopCat.showParentCat("option",shopCat.iParentCatID)%></select></td></tr><%end if
if shopCat.bShowOnlineCB then%><tr><td class=QSlabel><%=l("online")%></td><td><input type=checkbox name="bOnline" value="1" <%=convertChecked(shopCat.bOnline)%> /></td></tr><%end if%><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /><%if isNumeriek(shopCat.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><br /><br /><%end if%></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
