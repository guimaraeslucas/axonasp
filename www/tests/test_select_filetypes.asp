<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
Response.Buffer = False

Response.Write "=== TEST SELECT CASE WITH FILETYPES ===" & vbCrLf
Response.Flush

Dim fileType
fileType = "txt"

Response.Write "FileType: " & fileType & vbCrLf
Response.Flush

Response.Write "Starting Select Case..." & vbCrLf
Response.Flush

select case lcase(fileType)
    case "jpg","jpeg","jpe","jp2","jfif","gif","bmp","png","psd","eps","ico","tif","tiff","ai","raw","tga","mng","svg","doc","rtf","txt","wpd","wps","csv","xml","xsd","sql","pdf","xls","mdb","ppt","docx","xlsx","pptx","ppsx","artx","mp3","wma","mid","midi","mp4","mpg","mpeg","wav","ram","ra","avi","mov","flv","m4a","m4v","htm","html","css","swf","js","rar","zip","ogv","ogg","webm","tar","gz","eot","ttf","ics","woff","cod","msg","odt"
        Response.Write "MATCHED: File type is allowed" & vbCrLf
        Response.Flush
    case else
        Response.Write "NOT MATCHED: File type not allowed" & vbCrLf
        Response.Flush
end select

Response.Write "Select Case complete" & vbCrLf
Response.Flush

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
%>
