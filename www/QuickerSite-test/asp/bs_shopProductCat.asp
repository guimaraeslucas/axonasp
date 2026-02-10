<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bShoppingCart%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_qsc)%><!-- #include file="bs_shoppingcartmenu.asp"--><%dim shopProduct
set shopProduct=new cls_shopProduct
shopProduct.pick(decrypt(request("iShopProductID")))
dim catList,cat,subcats,subcat
set catList=new cls_shopCategory
set catList=catList.list
if request.form("btnaction")=l("save") then
shopProduct.saveCats(request.form("cat"))
message.Add("fb_saveOK")
end if
dim categories
set categories=shopProduct.categories%><form action="bs_shopProductCat.asp" method="post" name="mainform"><%=QS_secCodeHidden%><input type="hidden" name="iShopProductID" value="<%=encrypt(shopProduct.iID)%>" /><input type="hidden" value="<%=true%>" name="postback" /><table align=center cellpadding="4"><tr><td class="QSlabel">Product name:</td><td><%=sanitize(shopProduct.sName)%></td></tr><tr><td class=QSlabel valign="top">Select categories:</td><td align=left><ul style="margin-left:6px"><%for each cat in catList
response.write "<li style=""background-image:none""><input "
if categories.exists(convertGetal(cat)) then
response.write " checked=""checked"" "
end if
response.write " name=""cat"" type=""checkbox"" value=""" & encrypt(cat) & """ />" & sanitize(catList(cat).sName)
set subcats=catList(cat).subcategories
response.write "<ul>"
for each subcat in subcats
response.write "<li style=""background-image:none""><input "
if categories.exists(convertGetal(subcat)) then
response.write " checked=""checked"" "
end if
response.write " name=""cat"" type=""checkbox"" value=""" & encrypt(subcat) & """ />" & subcats(subcat).sName & "</li>"
next
set subcats=nothing
response.write "</ul></li>"
next
set catList=nothing%><ul></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /></td></tr></table></form><%set shopProduct=nothing%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
