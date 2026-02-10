<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><%dim customerList
set customerList=new cls_customerList
dim customers
set customers=customerList.table
dim c, urlApp
for each c in customers
urlApp=customers(c).sQSUrl & "/asp/ad_resetAPP.asp?apw="& sha256(C_ADMINPASSWORD)
if GetWebPage(urlApp) then
Response.Write customers(c).sUrl & " Reset OK"
else
Response.Write customers(c).sUrl & " <font color=Red><b>Reset NOT OK</b></font>"
end if
Response.Write "<br />"
Response.Flush 
next
function GetWebPage(ByVal Url)
on error resume next
GetWebPage=false
dim oXMLHTTP
Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
oXMLHTTP.open "GET", Url, false 
oXMLHTTP.send
if err.number <>0 then
err.Clear ()
exit function
end if
If oXMLHTTP.status=200 Then
if instr(oXMLHTTP.responseStream,"Application Reset")>0 then
GetWebPage=true
end if
End If
on error goto 0
End function
Response.End %>
