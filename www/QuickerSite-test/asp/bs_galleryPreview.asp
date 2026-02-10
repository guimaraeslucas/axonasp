<!-- #include file="begin.asp"-->
<%includeNS=true%>

<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bGallery%><%dim gallery,gDump
set gallery=new cls_gallery
gallery.backsitePV=true
gDump=treatconstants(gallery.build(),true)%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_gallery)%><table align=center><tr><td align=center>-> <b><a href="bs_galleryList.asp"><%=l("back")%></a></b> <-</td></tr></table><br /><br /><table width=600 align=center><tr><td><%=gDump%></td></tr><tr><td colspan=20 align="center" style="padding-top:10px"><input class="art-button" type="button" value="<%=l("modify")%>" onclick="javascript:location.assign('bs_galleryEdit.asp?iGalleryID=<%=encrypt(gallery.iId)%>')" /></td></tr></table><%set gallery=nothing%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
