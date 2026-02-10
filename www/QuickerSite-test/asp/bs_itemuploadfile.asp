<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Buffer				= true
Response.ContentType		= "text/html"
Response.CacheControl		= "no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires			= -1
Response.ExpiresAbsolute	= Now()-1
%>
<!-- #include file="assetmanager/security.asp"-->
<!-- #include file="asplite/asplite.asp"-->
<%

'on error resume next

dim uploadsDirVar, filesize, filename, sExt, fileKey

'default path: uploads
uploadsDirVar = server.MapPath (Application("QS_CMS_C_VIRT_DIR") & Application("QS_CMS_userfiles") & "pagemedia/" & aspl.convertNmbr(aspl.getRequest("iId"))) 

if not aspl.fso.folderexists(server.MapPath (Application("QS_CMS_C_VIRT_DIR") & Application("QS_CMS_userfiles") & "pagemedia/")) then
	aspl.fso.createfolder(server.MapPath (Application("QS_CMS_C_VIRT_DIR") & Application("QS_CMS_userfiles") & "pagemedia/"))
end if

if not aspl.fso.folderexists(uploadsDirVar) then
	aspl.fso.createfolder(uploadsDirVar)
end if

dim upload  : set upload = aspL.plugin("uploader")

upload.Save uploadsDirVar

for each fileKey in Upload.UploadedFiles.keys	

	filename	= Upload.UploadedFiles(fileKey).filename		
	sExt		= left(aspl.getFileType(filename),15)
	
	select case lcase(sExt)	
	
		case "jpg","jpeg","png"
		
		case else
			Upload.UploadedFiles(fileKey).delete
			err.raise 10,"File type","This filetype is not allowed"			
	
	end select
	
next

set upload=nothing

on error goto 0
%>