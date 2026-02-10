
<%sub logReferer()
if customer.bScanReferer then
dim Referer
Referer = Trim(Request.ServerVariables("HTTP_REFERER"))
if Referer<>"" then
if instr(1,Referer,customer.sUrl,vbTextCompare)=0 then
dim alternatedomains,ad
alternatedomains=split(convertStr(customer.sAlternateDomains),vbcrlf)
for ad=lbound(alternatedomains) to ubound(alternatedomains)
if instr(1,Referer,alternatedomains(ad),vbTextCompare)<>0 then exit sub
next
on error resume next
dim startpage
startpage=Trim(Request.ServerVariables("SCRIPT_NAME"))
if Request.ServerVariables("QUERY_STRING")<>"" then
startpage=startpage&"?"&Request.ServerVariables("QUERY_STRING")
end if
dim oSRS
set oSRS=db.getDynamicRS
oSRS.open("SELECT * FROM tblSession where 1=2")
oSRS.AddNew
oSRS("Browser") = Trim(Request.ServerVariables("HTTP_USER_AGENT"))
oSRS("sLanguage") = Trim(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE"))
oSRS("startpage") =  startpage
oSRS("Referer") =  Referer
oSRS("UserIP") = UserIP()
oSRS("iCustomerID") = cid
oSRS("dTS") = now()
oSRS.Update
oSRS.Close
Set oSRS = Nothing
session("extenalReferrer")=Referer
err.Clear ()
on error goto 0
end if
end if
end if
 
end sub
sub removeReferers()
dim rs
set rs=db.execute("delete from tblSession where iCustomerID="& cid)
set rs=nothing
end sub
function getReferersArr(dfrom,duntill,total)
dim rs,sql
set rs=db.getDynamicRS
sql = "select tblSession.referer,tblSession.browser,tblSession.startpage,tblSession.userip,tblSession.slanguage,tblSession.dTS  "
sql = sql & " from tblSession "
sql = sql & " where iCustomerID=" & cid
sql = sql & " and tblSession.dTs>="&getSQLDate(convertDateFromPicker(dfrom))
sql = sql & " and tblSession.dTs<="&getSQLDate(dateAdd("d",1,convertDateFromPicker(duntill)))
sql = sql & " order by tblSession.dTS desc"
rs.open(sql)
if not rs.eof then
getReferersArr=rs.getRows()
total=rs.recordCount
else
total=0
end if
set rs=nothing
end function%>
