<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bGallery%><!-- #include file="includes/header.asp"--><!-- #include file="includes/urlenCodeJS.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_gallery)%><%session("bHasSetUF")=false
session("showInsertButton")=false
dim gallery
set gallery=new cls_gallery
dim postBack
postBack=convertBool(Request.Form("postBack"))
if postBack then
gallery.getRequestValues()
end if
dim edID
edID=GeneratePassWord
select case Request.Form ("btnaction")
case l("save"), l("preview")
checkCSRF()
gallery.getRequestValues()
if gallery.save then
if Request.Form ("btnaction")=l("preview") then
Response.Redirect ("bs_gallerypreview.asp?iGalleryID=" & encrypt(gallery.iId))
else
Response.Redirect ("bs_galleryList.asp")
end if
end if
case l("delete")
checkCSRF()
gallery.remove
Response.Redirect ("bs_galleryList.asp")
end select
dim fe
set fe=new cls_fileexplorer
dim galleryTypeList
set galleryTypeList=new cls_galleryTypeList
dim gallerySEList
set gallerySEList=new cls_gallerySEList
dim galleryCycleList
set galleryCycleList=new cls_galleryCycleList%><!-- #include file="bs_galleryBack.asp"-->

<form action="bs_galleryEdit.asp" method="post" name="mainform">
<input type="hidden" name="iGalleryId" value="<%=encrypt(gallery.iID)%>" />
<input type="hidden" value="<%=true%>" name="postback" /><%=QS_secCodeHidden%>
<table align="center" cellpadding="2"><tr><td colspan="2" class="header"><%=l("general")%>:</td></tr>
<tr><td class="QSlabel"><%=l("name")%>:*</td><td><input type=text size=40 maxlength=50 name="sName" value="<%=quotRep(gallery.sName)%>" /></td></tr>
<tr><td class="QSlabel"><%=l("code")%>:*</td><td>[QS_GALLERY:<input type=text size=10 maxlength=45 name="sCode" value="<%=quotRep(gallery.sCode)%>" />]</td></tr>
<tr><td class="QSlabel"><%=l("pathtoimages")%>:*</td><td><select onchange="javascript:document.mainform.submit();" name="sPath"><%=fe.SelectBoxFolders(server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles")),gallery.sPath)%></select><%if not isLeeg(gallery.sPath) then%> <a class="QSPP" href="assetmanager/assetmanagerIF.asp?QS_inpFolder=<%=server.urlencode(server.mappath(replace(C_VIRT_DIR & Application("QS_CMS_userfiles") & gallery.sPath,"//","/",1,-1,1)))%>"><%=l("managepictures")%></a><%end if%></td></tr>
<tr><td class=QSlabel><%=l("selectgallerytype")%>:</td><td><select onchange="javascript:document.mainform.submit();" name="sType"><%=galleryTypeList.showSelected("option",gallery.sType)%></select></td></tr>
<tr><td class=QSlabel>Special effect:</td><td><select name="iSpecialEffect"><option value="0"></option><%=gallerySEList.showSelected("option",gallery.iSpecialEffect)%></select></td></tr>
<%
select case convertStr(gallery.sType)
	case QS_gallery_SS,QS_gallery_SC,QS_gallery_NS
	
	if convertStr(gallery.sType)=QS_gallery_SS then%>
		<tr><td class=QSlabel><%=l("max.imagesize")%>:</td><td><select name="iThumbSize"><%=numberList(10,1920,1,gallery.iThumbSize)%></select>&nbsp;px</td></tr>
		<tr><td class=QSlabel><%=l("imageborderwidth")%>:</td><td><select name="sBorder"><%=numberList(0,50,1,gallery.sBorder)%></select>&nbsp;px</td></tr>
		<tr><td class=QSlabel><%=l("imagebordercolor")%>:</td><td><input type="text" id="sBorderColor" name="sBorderColor" value="<%=quotrep(gallery.sBorderColor)%>" /><%=JQColorPicker("sBorderColor")%></td></tr>
	<%elseif convertStr(gallery.sType)=QS_gallery_SC then%>
		<tr><td class=QSlabel>Cycle effect:</td><td><select name="sCycleEffect"><%=galleryCycleList.showSelected("option",gallery.sCycleEffect)%></select></td></tr>
	<%elseif convertStr(gallery.sType)=QS_gallery_NS then%>
		<tr><td class=QSlabel>Cycle effect:</td><td><select name="sCycleEffect"><%=galleryCycleList.showSelectedNS2("option",gallery.sCycleEffect)%></select></td></tr>
		<tr><td class=QSlabel>Style:</td><td><select name="sNSCss"><%=galleryCycleList.showSelectedNS("option",gallery.sNSCss)%></select></td></tr>
		<tr><td class=QSlabel>DirectionNav?</td><td><input type="checkbox" name="bNSdirectionNav" value="1" <%=convertChecked(gallery.bNSdirectionNav)%> /></td></tr>
		<tr><td class=QSlabel>ControlNav?</td><td><input type="checkbox" name="bNSControlNav" value="1" <%=convertChecked(gallery.bNSControlNav)%> /></td></tr>
	<%end if%>
	
	<%
	if convertStr(gallery.sType)<>QS_gallery_NS then
	%>
	<tr><td class=QSlabel><%=l("fsr")%></td><td><input type="checkbox" name="bFSR" value="1" <%=convertChecked(gallery.bFSR)%> /></td></tr>
	<%
	end if
	%>
	<tr><td class=QSlabel>Sort images by:</td><td><select name="iSortImagesBy"><option <%if convertGetal(gallery.iSortImagesBy)=0 then response.write "selected='selected'"%> value="0">Name</option><option <%if convertGetal(gallery.iSortImagesBy)=4 then response.write "selected='selected'"%> value="4">Date created (recent first)</option><option <%if convertGetal(gallery.iSortImagesBy)=5 then response.write "selected='selected'"%> value="5">Date created (oldest first)</option><option <%if convertGetal(gallery.iSortImagesBy)=1 then response.write "selected='selected'"%> value="1">Date (recent first)</option><option <%if convertGetal(gallery.iSortImagesBy)=2 then response.write "selected='selected'"%> value="2">Date (oldest first)</option><option <%if convertGetal(gallery.iSortImagesBy)=3 then response.write "selected='selected'"%> value="3">Random sort</option></select></td></tr>
	<tr><td class=QSlabel><%=l("slideshowtimer")%>:</td><td><select name="iSlideShowTimerQS"><%=numberList(1,60,1,gallery.iSlideShowTimerQSDefault)%></select>&nbsp;<%=l("seconds")%></td></tr>
	
	<%if convertStr(gallery.sType)<>QS_gallery_NS then%>
	<tr><td class=QSlabel><%=l("slideshowwidth")%>:</td><td><select name="sWidth"><%=numberList(1,1920,1,gallery.sWidth)%></select>&nbsp;px</td></tr>
	<tr><td class=QSlabel><%=l("slideshowheight")%>:</td><td><select name="sHeight"><%=numberList(1,1200,1,gallery.sHeight)%></select>&nbsp;px</td></tr>
	<tr><td class="QSlabel"><%=l("customurlforimage")%></td><td><input type=text size=60 maxlength=255 name="sCustomLink" value="<%=quotRep(gallery.sCustomLink)%>" /></td></tr>
	<tr><td class=QSlabel><%=l("openinnewwindow")%></td><td><input name=bOpenInNewWindow type=checkbox value=1 <%if gallery.bOpenInNewWindow then Response.Write "checked"%> /></td></tr>
	<%
	else
	%>
	<tr><td class=QSlabel>Resize picture-width to:</td><td><select name="sWidth"><%=numberList(0,1920,1,gallery.sWidth)%></select>&nbsp;px (0: auto)</td></tr>
	<tr><td class=QSlabel>Resize picture-height to:</td><td><select name="sHeight"><%=numberList(0,1200,1,gallery.sHeight)%></select>&nbsp;px (0: auto)</td></tr>
	
	<%end if%>
	
	<%case else%>
<tr><td class=QSlabel><%=l("picsinrow")%>:</td><td><select name="iPicsInRow"><%=numberList(1,25,1,gallery.iPicsInRow)%></select></td></tr>
<tr><td class=QSlabel><%=l("browsepicsby")%>:</td><td><select name="iBrowseBy"><%=numberList(1,500,1,gallery.iBrowseBy)%></select></td></tr>
<tr><td class=QSlabel><%=l("thumbsize")%>:</td><td><select name="iThumbSize"><%=numberList(10,800,1,gallery.iThumbSize)%></select>&nbsp;pixels</td></tr>
<tr><td class=QSlabel><%=l("fsr")%></td><td><input type="checkbox" name="bFSR" value="1" <%=convertChecked(gallery.bFSR)%> /></td></tr>
<tr><td class=QSlabel><%=l("resizepicto")%>:</td><td><select name="iFullImageSize"><%=numberList(320,1920,16,gallery.iFullImageSize)%></select> px</td></tr>
<tr><td class=QSlabel><%=l("next")%>:</td><td><input type="text" size=40 maxlength="200" name="sNextLink" value="<%=sanitize(gallery.sNextLink)%>" /></td></tr>
<tr><td class=QSlabel><%=l("previous")%>:</td><td><input type="text" size=40 maxlength="200" name="sPreviousLink" value="<%=sanitize(gallery.sPreviousLink)%>" /></td></tr>
<tr><td class=QSlabel><%=l("expl_labelfullimage")%>:</td><td><input type="text" size=40 maxlength="200" name="sFullImage" value="<%=sanitize(gallery.sFullImage)%>" /></td></tr>
<tr><td class=QSlabel><%=l("showfilename")%></td><td><input type="checkbox" name="bShowFileName" value="1" <%=convertChecked(gallery.bShowFileName)%> /></td></tr>
<tr><td class=QSlabel>Sort images by:</td><td><select name="iSortImagesBy"><option <%if convertGetal(gallery.iSortImagesBy)=0 then response.write "selected='selected'"%> value="0">Name</option><option <%if convertGetal(gallery.iSortImagesBy)=4 then response.write "selected='selected'"%> value="4">Date created (recent first)</option><option <%if convertGetal(gallery.iSortImagesBy)=5 then response.write "selected='selected'"%> value="5">Date created (oldest first)</option><option <%if convertGetal(gallery.iSortImagesBy)=1 then response.write "selected='selected'"%> value="1">Date last modified (recent first)</option><option <%if convertGetal(gallery.iSortImagesBy)=2 then response.write "selected='selected'"%> value="2">Date last modified (oldest first)</option><option <%if convertGetal(gallery.iSortImagesBy)=3 then response.write "selected='selected'"%> value="3">Random sort</option></select></td></tr>
<tr><td class=QSlabel><%=l("autostartslideshow")%>?</td><td><input type="checkbox" name="bAutoStartSS" value="1" <%=convertChecked(gallery.bAutoStartSS)%> /></td></tr>
<tr><td class=QSlabel><%=l("slideshowtimer")%>:</td><td><select name="iSlideShowTimerQS"><%=numberList(1,60,1,gallery.iSlideShowTimerQSDefault)%></select>&nbsp;<%=l("seconds")%></td></tr>
<tr><td colspan=2 class=header><%=l("advanced")%>:</td></tr>
<tr><td class="QSlabel"><%=l("styletable")%>:</td><td><input type=text size=60 maxlength=255 name="sStyleTable" value="<%=quotRep(gallery.sStyleTable)%>" /></td></tr>
<tr><td class="QSlabel"><%=l("styletablecell")%>:</td><td><input type=text size=60 maxlength=255 name="sStyleTableCell" value="<%=quotRep(gallery.sStyleTableCell)%>" /></td></tr>
<tr><td class="QSlabel"><%=l("styleimage")%>:</td><td><input type=text size=60 maxlength=255 name="sStyleImage" value="<%=quotRep(gallery.sStyleImage)%>" /></td></tr>
<tr><td class="QSlabel"><%=l("customurlforimage")%></td><td><%=quotrep("<a ")%><input type=text size=60 maxlength=255 name="sCustomLink" value="<%=quotRep(gallery.sCustomLink)%>" /><%=quotrep(">")%><br /><small><%=l("variablesthatcanbeused:")%> [FILENAME], [COUNTER], [PAGEID]</small></td></tr><%end select%><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr>
<tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name="btnaction" value="<% =l("save")%>" /><%if isNumeriek(gallery.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><input class="art-button" type=submit name=btnaction value="<% =l("preview")%>" /><br /><br /><%end if%></td></tr></table>

<%if convertStr(gallery.sType)=QS_gallery_NS then%>
<p align="center">Manage captions/urls:</p>
<%'=gallery.sNSImgLinks%>
<table align="center" border="0" style="border-style:none">
<%
dim fso
set fso=server.createobject("scripting.filesystemobject")

dim imageFolder,imageFolderPath
imageFolderPath=C_VIRT_DIR & application("QS_CMS_userfiles") & gallery.sPath
imageFolderPath=replace(imageFolderPath,"//","/",1,-1,1)

if fso.folderexists(server.mappath(imageFolderPath)) then

    set imageFolder=fso.getFolder(server.mappath(imageFolderPath))

    dim sNSImgLinksARR,lineARR
    sNSImgLinksARR=split(gallery.sNSImgLinks,vbcrlf)


    dim file
    for each file in  imageFolder.files
    if allowedFileTypesforThumbing.exists(lcase(GetFileExtension(file.name))) then

	    'getValues
	    dim i,captionValue,urlValue,oinwValue
	    captionValue=""
	    urlValue=""
	    oinwValue=""
	    for i=lbound(sNSImgLinksARR) to ubound(sNSImgLinksARR)
		    if left(sNSImgLinksARR(i),len(removeDots(file.name)))=removeDots(file.name) then
			    lineArr=split(sNSImgLinksARR(i),vbtab)
			    captionValue=lineArr(1)
			    urlValue=lineArr(2)
			    on error resume next
			    if convertBool(lineArr(3)) then
				    oinwValue=" checked=""checked"" "
			    end if
			    on error goto 0
		    end if
	    next

    %>
	    <tr>
		    <td align="center"><img src="<%=C_DIRECTORY_QUICKERSITE%>/showthumb.aspx?maxSize=150&img=<%=imageFolderPath & "/" & file.name %>" /></td>
		    <td>
			    <table border="0" style="border-style:none">
			    <tr>
				    <td style="text-align:right">Caption:</td><td><input maxlength="1024" type="text" size="60" value="<%=sanitize(captionValue)%>" name="caption<%=removeDots(sanitize(file.name))%>" /></td>
			    </tr>
			    <tr>
				    <td style="text-align:right">Url:</td><td><input maxlength="1024" type="text" size="60" value="<%=sanitize(urlValue)%>" name="url<%=removeDots(sanitize(file.name))%>" /></td>
				    <td><input type="checkbox" name="oinw<%=removeDots(sanitize(file.name))%>" value="1" <%=oinwValue%>>new window</td>
			    </tr>
			    </table>
		    </td>
	    </tr>
	    <tr>
		    <td colspan=2><hr /></td>
	    </tr>
    <%
    end if
    next
end if
%>
</table>
<%
end if
%>
</form>
<%set fe=nothing



Response.Flush 
if convertGetal(gallery.iId)<>0 then
dim fsearch
set fsearch=new cls_fullSearch
fsearch.pattern="\[QS_GALLERY:+("&gallery.sCode&")+[\]]"
dim result
result=fsearch.search()
if not isLeeg(result) then
Response.Write "<p align=center>"&l("wheregalleryused")&"</p>"
Response.Write result
end if
set fsearch=nothing
end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
