
<%class cls_shopProduct
Public iId,bOnline,sName,dCreatedTS,dUpdatedTS,sLogo,sShortDesc,sLongDesc,iMakeID,iStock,sPriceA,sPriceB,sDefaultImage
Private Sub Class_Initialize
iId=null
bOnline=true
end sub
Private Sub Class_Terminate
end sub
Public property get allproducts
set allproducts=db.execute("select iId, sName, iMakeID from tblQShopProduct where iCustomerID=" & cId & " order by sName")
end property
Public Function Pick(id)
dim sql, RS, sValue
if isNumeriek(id) then
sql = "select * from tblQShopProduct where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sName	= rs("sName")
iStock	= rs("iStock")
sPriceA	= rs("sPriceA")
sPriceB	= rs("sPriceB")
sShortDesc	= rs("sShortDesc")
sLongDesc	= rs("sLongDesc")
iMakeID	= rs("iMakeID")
dCreatedTS	= rs("dCreatedTS")
dUpdatedTS	= rs("dUpdatedTS")
bOnline	= rs("bOnline")
sDefaultImage	= rs("sDefaultImage")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
exit function
end if
if not isLeeg(iStock) then
if not isNumeriek(iStock) then
check=false
message.AddError("err_mandatory")
exit function
end if
end if
End Function
Public Function Save
dim rs
if check() then
save=true
else
save=false
exit function
end if
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblQShopProduct where 1=2"
rs.AddNew
rs("dCreatedTS")	= now()
else
rs.Open "select * from tblQShopProduct where iId="& iId
end if
rs("sName")	= trim(left(sName,255))
rs("bOnline")	= convertBool(bOnline)
rs("sPriceA")	= sPriceA
rs("sPriceB")	= sPriceB
rs("sShortDesc")	= sShortDesc
rs("sLongDesc")	= sLongDesc
rs("iMakeID")	= convertGetal(iMakeID)
rs("iStock")	= convertGetal(iStock)
rs("iCustomerID")	= cId
rs("dUpdatedTS")	= now()
rs("iCustomerID")	= cId
rs("sDefaultImage")	= sDefaultImage
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function delete
dim fso
set fso=server.createobject("scripting.filesystemobject")
if fso.folderexists(server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & qsscart & "/" & shopProduct.iId)) then
fso.deleteFolder server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & qsscart & "/" & shopProduct.iId)
end if
db.execute("delete from tblQShopProduct where iId=" & convertGetal(iId))
end function
public function deleteImage(img)
dim fso
set fso=server.createobject("scripting.filesystemobject")
if fso.fileexists(server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & qsscart & "/" & shopProduct.iId & "/" & img)) then
fso.deletefile server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & qsscart & "/" & shopProduct.iId & "/" & img)
end if
set fso=nothing
if img=sDefaultImage then 
sDefaultImage=""
save()
end if
end function
Public function images
set images=server.createobject("scripting.dictionary")
dim fso
set fso=server.createobject("scripting.filesystemobject")
if fso.folderexists(server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & qsscart & "/" & iId)) then
dim files,file,folder
set folder=fso.getFolder(server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & qsscart & "/" & iId))
for each file in folder.files
images.add file.name,C_VIRT_DIR &  Application("QS_CMS_userfiles") & qsscart & "/" & iId & "/" & file.name
next
set files=nothing
end if
set fso=nothing
end function
Public function categories
set categories=server.createobject("scripting.dictionary")
dim rs,iCatID
set rs=db.execute("select iCategoryID from tblQShopProdCat where iProductID=" & convertGetal(iId))
while not rs.eof
iCatID=rs(0)
categories.Add iCatID,""
rs.movenext
wend
set rs=nothing
end function
public function setAsDefaultImage(img)
if not isLeeg(img) then
sDefaultImage=img
save()
end if
end function
public function saveCats(cats)
db.execute("delete from tblQShopProdCat where iProductID=" & convertGetal(iId))
dim arrC,i
arrC=split(cats,",")
for i=lbound(arrC) to ubound(arrC)
if convertGetal(decrypt(trim(arrC(i))))<>0 then
db.execute("insert into tblQShopProdCat (iProductID,iCategoryID) values(" & iId & "," & convertGetal(decrypt(trim(arrC(i)))) & ")")
end if
next
end function
end class%>
