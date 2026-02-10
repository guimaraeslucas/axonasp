<!--#include file="freeaspupload.asp"-->
<%
if not session("authenticated") then Response.End 

Dim uploadsDirVar
uploadsDirVar = server.MapPath ("files") 

dim Upload
Set Upload = New FreeASPUpload
Upload.Save uploadsDirVar

dim ks, fileKey, strMessage
ks = Upload.UploadedFiles.keys

if (UBound(ks) <> -1) then

	for each fileKey in Upload.UploadedFiles.keys
	
		select case lcase(GetFileExtension(Upload.UploadedFiles(fileKey).FileName))
		
			case "mdb","xls","txt","htm","html"
			
				if lcase(Upload.UploadedFiles(fileKey).Length)>15000000 then
					Upload.UploadedFiles(fileKey).delete()
					Response.Redirect ("upload.asp?message=filetobig")
				end if	
				
				Response.Redirect ("upload.asp?message=ok")
			
			case else
				
				Upload.UploadedFiles(fileKey).delete()
				Response.Redirect ("upload.asp?message=filetype")
		
		end select
	
	next
	
else
	Response.Redirect ("upload.asp?message=nofile")
end if



%>