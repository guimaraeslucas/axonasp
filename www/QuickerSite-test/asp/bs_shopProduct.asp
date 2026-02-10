<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bShoppingCart%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_qsc)%><!-- #include file="bs_shoppingcartmenu.asp"--><%dim shopProduct
set shopProduct=new cls_shopProduct
shopProduct.pick(decrypt(request("iShopProductID")))
dim showMake
set showMake=new cls_shopMake
dim postback
postback=convertBool(request.form("postback"))
if postback then
select case request.form("btnaction")
case l("save")
shopProduct.sName=request.form("sName")
shopProduct.bOnline=request.form("bOnline")
shopProduct.sShortDesc=request.form("sShortDesc")
shopProduct.sLongDesc=request.form("sLongDesc")
shopProduct.iStock=request.form("iStock")
shopProduct.iMakeID=request.form("iMakeID")
if shopProduct.save then response.redirect ("bs_shoppingcart.asp")
case l("delete")
shopProduct.delete
response.redirect ("bs_shoppingcart.asp")
end select
end if%><form action="bs_shopProduct.asp" method="post" name="mainform"><%=QS_secCodeHidden%><input type="hidden" name="iShopProductID" value="<%=encrypt(shopProduct.iID)%>" /><input type="hidden" value="<%=true%>" name="postback" /><table align=center cellpadding="4"><tr><td class="QSlabel">Product name:*</td><td><input type=text size=40 maxlength=250 name="sName" value="<%=sanitize(shopProduct.sName)%>" /></td></tr><tr><td class="QSlabel">Make</td><td><select name="iMakeID"><option value=""></option><%=showMake.showMake("option",shopProduct.iMakeID)%></select></td></tr><tr><td class="QSlabel">Short Description</td><td><textarea rows=3 cols=60 name="sShortDesc"><%=sanitize(shopProduct.sShortDesc)%></textarea></td></tr><tr><td class="QSlabel">Long Description</td><td><%createFCKInstance shopProduct.sLongDesc,"siteBuilderRichText","sLongDesc"%></td></tr><tr><td class="QSlabel">Stock</td><td><input type=text size=8 maxlength=10 name="iStock" value="<%=sanitize(shopProduct.iStock)%>" /></td></tr><tr><td class=QSlabel><%=l("online")%></td><td><input type=checkbox name="bOnline" value="1" <%=convertChecked(shopProduct.bOnline)%> /></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /><%if isNumeriek(shopProduct.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><br /><br /><%end if%></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
