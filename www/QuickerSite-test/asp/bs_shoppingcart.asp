<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bShoppingCart%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_qsc)%><!-- #include file="bs_shoppingcartmenu.asp"--><%dim catList,cat,subcats,subcat,shopMake,shopMakeList
set catList=new cls_shopCategory
set catList=catList.list
set shopMake=new cls_shopMake
set shopMakeList=shopMake.list%><table style="border-style:none" cellpadding=5><tr><td valign="top"><blockquote style="background-image:none;padding-left:10px"><h5>Categories:</h5><ul><%for each cat in catList
response.write "<li><strong><a href='bs_shopAdCat.asp?iShopCatID=" & encrypt(cat) & "'>" & sanitize(catList(cat).sName) & "</a></strong>"
if not convertBool(catList(cat).bOnline) then response.write "&nbsp;<span style=""color:" & MYQS_offlineLinkColor & """><i>(offline)</i></span>"
set subcats=catList(cat).subcategories
response.write "<ul>"
for each subcat in subcats
response.write "<li><a href='bs_shopAdCat.asp?iShopCatID=" & encrypt(subcat) & "'>" & subcats(subcat).sName & "</a>"
if not convertBool(subcats(subcat).bOnline) then response.write "&nbsp;<span style=""color:" & MYQS_offlineLinkColor & """><i>(offline)</i></span>"
response.write "</li>"
next
set subcats=nothing
response.write "</ul></li>"
next%></ul><span class="art-button-wrapper"><span class="art-button-l"> </span><span class="art-button-r"> </span><a class="art-button" href="bs_shopAdCat.asp">Add Category</a></span>&nbsp;&nbsp;
</blockquote></td><td valign="top"><blockquote style="background-image:none;padding-left:10px"><h5>Makes:</h5><ul><%for each cat in shopMakeList
response.write "<li><strong><a href='bs_shopMake.asp?iShopMakeID=" & encrypt(cat) & "'>" & sanitize(shopMakeList(cat).sName) & "</a></strong>"
if not convertBool(shopMakeList(cat).bOnline) then response.write "&nbsp;<span style=""color:" & MYQS_offlineLinkColor & """><i>(offline)</i></span>"
response.write "</li>"
next%></ul><span class="art-button-wrapper"><span class="art-button-l"> </span><span class="art-button-r"> </span><a class="art-button" href="bs_shopMake.asp">Add Make</a></span>&nbsp;&nbsp;
</blockquote></td><%set catList=nothing
dim shopProduct,encID
set shopProduct=new cls_shopProduct
dim allproducts,flatlist
set allproducts=shopProduct.allproducts%><td valign="top"><blockquote style="background-image:none;padding-left:10px"><h5>Products:</h5><table cellpadding=4 cellspacing=0><%while not allproducts.eof
encID=encrypt(allproducts(0))
response.write "<tr>"
response.write "<td style=""border-bottom:1px solid #DDDDDD""><a href=""bs_shopProduct.asp?iShopProductID=" & encID & """>" & sanitize(allproducts(1)) & "</a></td>"
response.write "<td style=""border-bottom:1px solid #DDDDDD"">"
if shopMakeList.exists(convertGetal(allproducts(2)))<>0 then
response.write sanitize(shopMakeList(convertGetal(allproducts(2))).sName)
end if
response.write "</td>"
response.write "<td style=""border-bottom:1px solid #DDDDDD""><span class=""art-button-wrapper""><span class=""art-button-l""> </span><span class=""art-button-r""> </span><a class=""art-button"" href=""bs_shopProduct.asp?iShopProductID=" & encID & """>Edit</a></span></td>"
response.write "<td style=""border-bottom:1px solid #DDDDDD""><span class=""art-button-wrapper""><span class=""art-button-l""> </span><span class=""art-button-r""> </span><a class=""art-button"" href=""bs_shopProductCat.asp?iShopProductID=" & encID & """>Categories</a></span></td>"
response.write "<td style=""border-bottom:1px solid #DDDDDD""><span class=""art-button-wrapper""><span class=""art-button-l""> </span><span class=""art-button-r""> </span><a class=""art-button"" href=""bs_shopProductImg.asp?iShopProductID=" & encID & """>Images</a></span></td>"
response.write "</tr>"
allproducts.movenext
wend%></table><span class="art-button-wrapper"><span class="art-button-l"> </span><span class="art-button-r"> </span><a class="art-button" href="bs_shopProduct.asp">Add Product</a></span>&nbsp;&nbsp;
</blockquote></td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
