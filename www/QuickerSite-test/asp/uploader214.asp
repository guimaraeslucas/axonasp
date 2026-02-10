
<!-- #include file="uploaderCLS.asp" --><%'security
on error resume next
if cstr(application("UserIP"))<>cstr(UserIP) then response.end 
if application("mupath")="" then response.end 
if application("sessionID")<>request.querystring("sId") then response.end
if application("sessionID")="" then response.end
if request.querystring("sId")="" then response.end
if err.number<>0 then response.end 
on error goto 0
dim UploadiFyPath,UploadiFyFolder,UploadifyObject
UploadiFyPath = Request.ServerVariables("PATH_TRANSLATED")
UploadiFyPath = Replace(UploadiFyPath,"uploader214.asp","",1,-1,1)
UploadiFyFolder=application("mupath")
Set UploadifyObject = New Uploader
UploadifyObject.Save(Server.MapPath(UploadiFyFolder))
Response.Write("<HTML><HEAD></HEAD><BODY></BODY></HTML>")
Function UserIP()
' This returns the True IP of the client calling the requested page
' Checks to see if HTTP_X_FORWARDED_FOR has a value then the client is operating via a proxy
UserIP = Server.HTMLEncode(Server.URLEncode(Request.ServerVariables ( "HTTP_X_FORWARDED_FOR" )))
If UserIP = "" Then
UserIP = Request.ServerVariables ( "REMOTE_ADDR" )
End if
if isNull(UserIP) or UserIP="" then response.end
End Function%>
