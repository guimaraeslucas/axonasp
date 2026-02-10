<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bGallery%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_gallery)%><%dim copyGallery
set copyGallery=new cls_gallery
copyGallery.copy()
if convertGetal(copyGallery.iId)<>0 then Response.Redirect ("bs_galleryEdit.asp?iGalleryID="&encrypt(copyGallery.iId))
dim galleries
set galleries=customer.galleries%><p align=center><%=getArtLink("bs_galleryEdit.asp",l("newgallery"),"","","")%></p><%if galleries.count>0 then%><table align=center cellspacing=0 cellpadding=4><%dim galleryKey
for each galleryKey in galleries%><tr><td style="border-top:1px solid #DDD"><a href="bs_galleryEdit.asp?iGalleryID=<%=encrypt(galleryKey)%>"><%=galleries(galleryKey).sName%></a></td><td style="border-top:1px solid #DDD"><%=getIcon(l("preview"),"search","bs_galleryPreview.asp?iGalleryID="& encrypt(galleryKey),"","preview"&galleryKey)%></td><td style="border-top:1px solid #DDD"><%=getIcon(l("copyItem"),"copyItem","bs_galleryList.asp?"&QS_secCodeURL&"&amp;iGalleryID="&encrypt(galleryKey),"javascript:return confirm('"& l("areyousuretocopy") &"');","copy"&galleryKey)%></td><td style="border-top:1px solid #DDD"><input size="20" type="text" onclick="javascript:this.select();" value="[QS_GALLERY:<%=sanitize(galleries(galleryKey).sCode)%>]" /></td><td style="border-top:1px solid #DDD"><i>iID: <%=galleryKey%></i></td></tr><%next%></table><%else%><p align=center><%=l("nogallery")%></p><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
