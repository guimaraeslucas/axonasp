<!-- #include file="begin.asp"-->
<%if request.querystring("I")=sha256(cId) then
dim fso1,sizeFOlder
set fso1=server.createobject("scripting.filesystemobject")
set sizeFOlder=fso1.getfolder(server.mappath(C_VIRT_DIR & Application("QS_CMS_userfiles")))
response.write sizeFOlder.size
set fso1=nothing
end if%>
