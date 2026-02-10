
<%class cls_catalogItemFile
Public iId
Public sTitle
Public dUpdatedTS
Public dCreatedTS
Public sName
Public iFileTypeID
Public iItemID

private p_catalogitem

Private Sub Class_Initialize
	set p_catalogitem=nothing
end sub

public function catalogitem
	if p_catalogitem is nothing then
		set p_catalogitem= new cls_catalogItem
		p_catalogitem.pick(iItemID)
	end if
	set catalogitem=p_catalogitem
end function


Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblCatalogItemFiles where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
iItemID	= rs("iItemID")
sTitle	= rs("sTitle")
sName	= rs("sName")
iFileTypeID	= rs("iFileTypeID")
dUpdatedTS	= rs("dUpdatedTS")
dCreatedTS	= rs("dCreatedTS")
end if
set RS = nothing
end if
end function
Public property get url

if allowedFileTypesforThumbing.exists(GetFileExtension(sName)) then
	url="<a class=""QSPPIMG"" href=""" & C_DIRECTORY_QUICKERSITE & "/showthumb.aspx?maxsize=740&img=" & C_VIRT_DIR & Application("QS_CMS_userfiles") & catalogitem.catalog.correFP & sName &"""><img style=""box-shadow:1px 1px 3px #555"" src=""" & C_DIRECTORY_QUICKERSITE & "/showthumb.aspx?maxsize=150&img=" & C_VIRT_DIR & Application("QS_CMS_userfiles") & catalogitem.catalog.correFP & sName &  """ /></a>"
else
	url="<a target='" & encrypt(iId) & "' href=" & C_VIRT_DIR & Application("QS_CMS_userfiles") & catalogitem.catalog.correFP & sName &">" & catalogItem.catalog.showSelectedFileType("single",encrypt(iFileTypeID)) & "</a>"
end if

end property
Public property get simpleurl
simpleurl=C_VIRT_DIR & Application("QS_CMS_userfiles") & catalogitem.catalog.correFP & sName
end property

Public Function Save()
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblCatalogItemFiles where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblCatalogItemFiles where iId="& iId
end if
rs("sTitle")	= sTitle
rs("sName")	= sName
rs("dUpdatedTS")	= now()
rs("iFileTypeID")	= iFileTypeID
rs("iItemID")	= iItemID
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public sub remove
if isNumeriek(iId) then
db.execute("delete from tblCatalogItemFiles where iId="& convertGetal(iId))
dim fso
set fso=server.CreateObject ("scripting.filesystemobject")
if fso.FileExists (server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & catalogitem.catalog.correFP & sName)) then
fso.DeleteFile server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & catalogitem.catalog.correFP & sName)
end if
set fso=nothing
end if
end sub
public function copy(itemID)
dim pw, sNewName
pw=GeneratePassWord()
iItemID=itemID
iId=null
Save()
dim fso2
set fso2=server.CreateObject ("scripting.filesystemobject")
if fso2.FileExists (server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & catalogitem.catalog.correFP & sName)) then
sNewName=encrypt(iId) & "_" & pw & "." & GetFileExtension(sName)
fso2.copyFile server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles")  & catalogitem.catalog.correFP & sName),server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles") & catalogitem.catalog.correFP & sNewName)
sName=sNewName
save()
end if
set fso2=nothing
end function
end class%>
