<!-- #include file="begin.asp"-->
<%
pagetoemail=true
editPage=true
dim getLperP, iLPID, page
set page=new cls_page
page.pick(decrypt(request("iId")))

'FSO
dim fso
set fso=server.createobject("scripting.filesystemobject")
if not fso.folderExists(server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & "listitemimages")) then
	fso.createfolder server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & "listitemimages") 
end if
set fso=nothing


dim  ks, fileKey, strMessage, fbM
Dim uploadsDirVar
uploadsDirVar = server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & "listitemimages") 

dim Upload
Set Upload = New FreeASPUpload

if isLeeg(page.iId) then

	Upload.Save uploadsDirVar
	checkCSRF_Upload(Upload.form("QSSEC"))
	
	page.pick(decrypt(upload.form("iId")))
	
end if

set getLperP=logon.contact.getLper
if not isLeeg(page.iListPageID) then
	iLPID	= convertGetal(page.iListPageID)
	selectedPage.iListPageID	= iLPID
else
	iLPID=convertGetal(selectedPage.iListPageID)
end if

if not getLperP.exists(iLPID) then Response.Redirect (C_DIRECTORY_QUICKERSITE & "/asp/noaccess.htm")

if isLeeg(page.sItemPicture) then	

	ks = Upload.UploadedFiles.keys

	if isNumeriek(decrypt(upload.form("iId"))) then
		page.pick(decrypt(upload.form("iId")))
	end if

	if (UBound(ks) <> -1) then

		for each fileKey in Upload.UploadedFiles.keys
			
			select case lcase(GetFileExtension(Upload.UploadedFiles(fileKey).FileName)) 
				
				case "png","jpg","gif","jpeg"
				
				case else
					strMessage=strMessage&"Only JPG/GIF/PNG files please."
					
			end select		

			if not isLeeg(strMessage) then 
				Upload.UploadedFiles(fileKey).delete()
			else
				'hernoemen naar page.iID			
				Upload.UploadedFiles(fileKey).rename page.iId & "." & GetFileExtension(Upload.UploadedFiles(fileKey).FileName),uploadsDirVar
				'nog bewaren in de pagina!
				page.sItemPicture=GetFileExtension(Upload.UploadedFiles(fileKey).FileName)
				page.sLPIC="fp"
				page.save()		
				response.redirect ("fs_editListItemPic.asp?iId=" & encrypt(page.iId))		
			end if
		next

	end if

	if not isLeeg(strMessage) then
		strMessage="<p style=""color:Red""><strong>" & strMessage & "</strong></p>"
	end if

else

	if request.form("btnAction")<>"" then
		page.sLPIC=request.form("sLPIC")
		page.save()	
		response.redirect ("fs_editListItems.asp?iId=" & encrypt(iLPID))
	end if

	if request.form("delAction")<>"" then
		page.deleteListItemImage()
		response.redirect ("fs_editListItems.asp?iId=" & encrypt(iLPID))
	end if

end if

%>
<!-- #include file="includes/commonheader.asp"-->
<body style="color:#000;background-color:#FFF;background-image:none" onload="javascript:window.focus();">
<%=strMessage%>
<%
if isLeeg(page.sItemPicture) then
%>
<div style="margin:0 auto;width:100px;text-align:center"><br /><a href="fs_editListItems.asp?iId=<%=encrypt(iLPID)%>" class="art-button"><%=l("back")%></a><br /><br /></div>
<form method="post" enctype="multipart/form-data" action="fs_editListItemPic.asp" id=upload name=upload>
<%=QS_secCodeHidden%>
<input type="hidden" value="<% =EnCrypt(page.iId) %>" name="iId" />
<table align="center" cellspacing="0" cellpadding="3">
	<tr>
		<td><INPUT TYPE="FILE" onchange="javascript: document.upload.uploadButton.value='please wait...';document.upload.submit();" NAME="picture" /></td>
		<td><INPUT class="art-button" TYPE=SUBMIT VALUE="<%=l("upload")%>" id=SUBMIT1 name="uploadButton" /></td>
	</tr>
</table>
</form>
<%
else
%>
<div style="margin:0 auto;width:270px;text-align:center"><img style="margin:20px 0px 20px 0px" src="<%=C_DIRECTORY_QUICKERSITE%>/showthumb.aspx?a=<%=generatePassword%>&amp;maxSize=270&amp;img=<%=C_VIRT_DIR &  Application("QS_CMS_userfiles") & "listitemimages/" & page.iId & "." & page.sItemPicture%>" alt="" /></div>

<form method="post" action="fs_editListItemPic.asp" id=editP name=editP>
<%=QS_secCodeHidden%>
<input  type="hidden" value="<% =EnCrypt(page.iId) %>" name="iId" />
<table align="center" cellspacing="0" cellpadding="3">
	<tr>
		<td><input type="radio" name="sLPIC" <%if page.sLPIC="al" then response.write " checked=""checked"" "%> value="al" /></td><td>Align left (50%)</td>
	</tr>
	<tr>
		<td><input type="radio" name="sLPIC" <%if page.sLPIC="fp" then response.write " checked=""checked"" "%> <%if isLeeg(page.sLPIC) then response.write " checked=""checked"" "%> value="fp" /></td><td>Full page width (100%)</td>
	</tr>
	<tr>
		<td><input type="radio" name="sLPIC" <%if page.sLPIC="ar" then response.write " checked=""checked"" "%> value="ar" /></td><td>Align right (50%)</td>
	</tr>
	<tr>
		<td><input type="radio" name="sLPIC" <%if page.sLPIC="CC" then response.write " checked=""checked"" "%> value="CC" /></td><td>custom CSS on class="ListItemPictureCSS"</td>
	</tr>
	<tr>
		<td colspan="2"><input class="art-button" type=submit value="Save" name="btnAction" />  <input class="art-button" type=submit onclick="javascript:return confirm('Are you sure to delete the picture?');" value="Delete Picture" name="delAction" /></td>
	</tr>
</table>
</form>
<%
end if
%>
</body>
</html>
<%cleanUPASP%>