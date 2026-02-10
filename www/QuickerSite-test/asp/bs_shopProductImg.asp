<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bShoppingCart%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_qsc)%><!-- #include file="bs_shoppingcartmenu.asp"--><%dim shopProduct,images,image
set shopProduct=new cls_shopProduct
shopProduct.pick(decrypt(request("iShopProductID")))
'delete image
if not isLeeg(request.querystring("delete")) then
shopProduct.deleteImage(request.querystring("delete"))
end if
if not isLeeg(request.querystring("default")) then
shopProduct.setAsDefaultImage(request.querystring("default"))
end if
set images=shopProduct.images
for each image in images%><div style="float:left;text-align:center;margin:5px"><a href="<%=C_DIRECTORY_QUICKERSITE%>/showthumb.aspx?img=<%=images(image)%>&amp;maxsize=600" class="QSPPIMG"><img style="margin-bottom:3px" width="120" src="<%=C_DIRECTORY_QUICKERSITE%>/showthumb.aspx?img=<%=images(image)%>&amp;maxsize=150&amp;FSR=1" alt="<%=image%>" /></a><div class="cleared"></div><%if image=shopProduct.sDefaultImage then%><span><i>default image</i></span><%else%><a href="bs_shopProductImg.asp?default=<%=server.urlencode(image)%>&iShopProductID=<%=encrypt(shopProduct.iId)%>">set as default image</a><%end if%><div class="cleared"></div><span style="margin-top:3px" class="art-button-wrapper"><span class="art-button-l"> </span><span class="art-button-r"> </span><a onclick="javascript: return confirm('<%=l("areyousure")%>');" class="art-button" href="bs_shopProductImg.asp?delete=<%=server.urlencode(image)%>&iShopProductID=<%=encrypt(shopProduct.iId)%>">Delete</a></span></div><%next%><div class="cleared"></div><form ENCTYPE="multipart/form-data" action="bs_shopProductImgUpload.asp?iShopProductID=<%=encrypt(shopProduct.iId)%>" method="post" name="mainform"><%=QS_secCodeHidden%><input type="hidden" name="iShopProductID" value="<%=encrypt(shopProduct.iID)%>" /><input type="hidden" value="<%=true%>" name="postback" /><table align=center cellpadding="4"><tr><td class="QSlabel">Product name:</td><td><%=sanitize(shopProduct.sName)%></td></tr><tr><td class=QSlabel>Select an image:</td><td align=left><INPUT TYPE="FILE" NAME="image1" /> (only jpg files)</td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("upload")%>" /></td></tr></table></form><%set shopProduct=nothing%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
