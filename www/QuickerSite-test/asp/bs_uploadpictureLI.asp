<!-- #include file="begin.asp"-->
<!-- #include file="bs_security.asp"-->
<%

if not convertBool(customer.bListItemPic) then response.end 

dim page
set page=new cls_page
page.pick(decrypt(request.querystring("iId")))

dim fso
set fso=server.createobject("scripting.filesystemobject")
if not fso.folderExists(server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & "listitemimages")) then
	fso.createfolder server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & "listitemimages") 
end if
set fso=nothing

Dim uploadsDirVar
uploadsDirVar = server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & "listitemimages") 

dim Upload
Set Upload = New FreeASPUpload

Upload.Save uploadsDirVar
checkCSRF_Upload(Upload.form("QSSEC"))

dim ks, fileKey, strMessage, fbM
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
			response.redirect ("bs_editpictureLI.asp?iId=" & encrypt(page.iId))		
		end if
	next

end if

if not isLeeg(strMessage) then
	strMessage="<p style=""color:Red""><strong>" & strMessage & "</strong></p>"
end if

%>

<html>
<head>
<title>Upload Picture</title>
<%=fbM%>
</head>
<body style="background-color:#FFFFFF">
<%=strMessage%>
<form method="post" enctype="multipart/form-data" action="bs_uploadpictureLI.asp" id=upload name=upload>
<%=QS_secCodeHidden%>
<input type="hidden" value="<% =EnCrypt(page.iId) %>" name="iId" />
<table>
	<tr>
		<td><INPUT TYPE="FILE" onchange="javascript: document.upload.uploadButton.value='please wait...';document.upload.submit();" NAME="picture" /></td>
		<td><INPUT class="art-button" TYPE=SUBMIT VALUE="<%=l("upload")%>" id=SUBMIT1 name="uploadButton" /></td>
	</tr>
</table>
</form>
</body>
</html>
<%cleanUPASP%>
