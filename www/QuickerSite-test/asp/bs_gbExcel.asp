<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%dim rs
set rs=db.execute("select * from tblGuestBookItem where iGuestbookID=" & convertGetal(decrypt(request.querystring("iGBID"))) & " order by iId asc")
dim gbTable
gbTable="<table>"
while not rs.eof
gbTable=gbTable & "<tr>"
gbTable=gbTable & "<td>" & rs(0) & "</td>"
gbTable=gbTable & "<td>" & rs(1) & "</td>"
gbTable=gbTable & "<td>" & rs(2) & "</td>"
gbTable=gbTable & "<td>" & rs(3) & "</td>" 
gbTable=gbTable & "<td>" & rs(4) & "</td>" 
gbTable=gbTable & "<td>" & rs(5) & "</td>"
gbTable=gbTable & "</tr>"
rs.movenext
wend
gbTable=gbTable & "</table>"
set rs=nothing
dim excelfile
set excelfile=new cls_excelfile
excelfile.export(gbTable)
excelfile.redirectLink()
set excelfile=nothing %>
