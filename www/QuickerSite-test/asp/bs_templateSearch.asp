<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bTemplates%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Setup)%><%=getBOSetupMenu(btn_template)%><%dim browseBy,arr
browseBy=30
if request.querystring("sBrowseOnlineTemplatesUrl")<>"" then
sBrowseOnlineTemplatesUrl=request.querystring("sBrowseOnlineTemplatesUrl")
end if
dim oXMLHTTP,arrF
if not isLeeg(request.querystring("install")) then
Set oXMLHTTP = server.createobject("MSXML2.ServerXMLHTTP")
'ServerXMLHTTP
oXMLHTTP.open "GET", sBrowseOnlineTemplatesUrl & "list.asp?version=" & C_QS_VERSION & "&install=" & server.urlencode(request.querystring("install"))
oXMLHTTP.send
arr=split(oXMLHTTP.responseText,vbcrlf)
dim fso,arrG,install
install=request.querystring("install")
install=replace(install,sBrowseOnlineTemplatesUrl,"",1,-1,1)
arrG=split(install,"/")
set fso=server.createObject("scripting.filesystemobject")
on error resume next
fso.createfolder(server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates"))
fso.createfolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "\" & arrG(0)
fso.createfolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "\" & arrG(0) & "\" & arrG(1)
fso.createfolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "\" & arrG(0) & "\" & arrG(1) & "\images"
on error goto 0
for arrF=lbound(arr) to ubound(arr)
if not isLeeg(arr(arrF)) then
oXMLHTTP.open "GET", sBrowseOnlineTemplatesUrl & arr(arrF)
oXMLHTTP.send
if oXMLHTTP.status=200 then
SaveBinaryData server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates/") & "\" & replace(arr(arrF),"/","\",1,-1,1), oXMLHTTP.responseBody
end if
end if
next
set oXMLHTTP=nothing
set fso=nothing
response.redirect("bs_createTemplate.asp?path=" & server.urlencode(server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "/" & replace(install,"sc.png","",1,-1,1)))
end if
if isLeeg(session("fullListTemplates")) then
Set oXMLHTTP = server.createobject("MSXML2.ServerXMLHTTP")
oXMLHTTP.open "GET", sBrowseOnlineTemplatesUrl & "list.asp?v=" & C_QS_VERSION,false
oXMLHTTP.send
session("fullListTemplates")=oXMLHTTP.responseText
arr=split(oXMLHTTP.responseText,vbcrlf)
set oXMLHTTP=nothing
else
arr=split(session("fullListTemplates"),vbcrlf)
end if
dim iStart
iStart=convertGetal(request.querystring("iStart"))
if iStart=0 then iStart=1%><table cellpadding="1" cellspacing="1" style="border-style:none;text-align:center"><%=navbalk%><tr><td><%dim i,count,rows
count=0
i=0
for i=lbound(arr) to ubound(arr)
if not isLeeg(arr(i)) then
if iStart*browseBy<i+browseBy+1  then
rows=rows& "<div style=""margin:1px;float:left;padding:3px;border-bottom:1px solid #DDD;text-align:center"">"
rows=rows& "<a title=""Preview!"" class=""bPopupFullWidthNoReload"" href=""" & replace(arr(i),"sc.png","page.html",1,-1,1) & """><img style=""width:180px;height:180px;border-style:none;margin:1px"" border=""0"" src=""" & arr(i) & """ /></a>"
rows=rows& "<br /><a onclick=""javascript:return confirm('Are you sure to download and install this template?');"" href=""bs_templateSearch.asp?install=" & server.urlencode(arr(i)) & """>Download &amp; Install</a></div>"
count=count+1
if count=browseBy then exit for
end if
end if
next
response.write rows%></td></tr><%=navbalk%></table><%function navbalk%><tr><td colspan=5 style="text-align:center"><%dim max
max=ubound(arr)
dim oo
oo=1
while (oo*browseBy)-browseBy<max
if convertGetal(iStart)=oo then
response.write "<div style=""text-align:center;vertical-align:middle;border:1px solid #BBB;float:left;padding:1px;margin:1px;background-color:#FFF;width:18px""><a style=""font-weight:700"" href=""bs_templateSearch.asp?iStart=" & oo & """>" 
else
response.write "<div style=""text-align:center;vertical-align:middle;border:1px solid #BBB;float:left;padding:1px;margin:1px;background-color:#EEE;width:18px""><a href=""bs_templateSearch.asp?iStart=" & oo & """>" 
end if
response.write oo & "</a></div>"
oo=oo+1
wend%></td></tr><%end function%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
