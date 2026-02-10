<!-- #include file="begin.asp"-->
<!-- #include file="bs_security.asp"-->
<%logon.hasaccess secondAdmin.bTemplates%>
<!-- #include file="includes/header.asp"-->
<!-- #include file="bs_initBack.asp"-->
<!-- #include file="bs_header.asp"-->
<%=getBOHeader(btn_Setup)%>
<%=getBOSetupMenu(btn_template)%>
<!-- #include file="bs_templateBack.asp"-->
<%
dim postBack
postBack=convertBool(Request.Form("postBack"))
Dim zipper
set zipper=new cls_zipper
if not customer.supportZipper then response.redirect("bs_templateList.asp")
dim defaultFolder
defaultFolder="/templates"
dim fe,bRemove,bContinue
set fe=new cls_fileexplorer
bRemove=true
bContinue=true
if postBack then
defaultFolder=Request.Form ("sPath")
dim foldername
foldername=Request.Form ("foldername")
if isLeeg(foldername) then
foldername=server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles")) & replace(defaultFolder,"/","\",1,-1,1)  & "\" & replace(Request.Form ("zip"),".zip","",1,-1,1)
else
foldername=server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles")) & replace(defaultFolder & "/" & replace(foldername,"/","",1,-1,1),"/","\",1,-1,1)
end if
zipper.zipname=C_VIRT_DIR & Application("QS_CMS_userfiles") & Request.Form ("zip")
zipper.ziponly=Request.Form ("zip")
'unpack
if zipper.unpack (server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles")) & "\" & Request.Form ("zip"),foldername) then
'remove zip?
if convertBool(Request.form("bRemove")) then
'if fe.fileexists(C_VIRT_DIR & Application("QS_CMS_userfiles") & Request.Form ("zip")) then
fe.deleteFile C_VIRT_DIR & Application("QS_CMS_userfiles") & Request.Form ("zip")
'end if
end if
'open folder
if not message.hasErrors then
if zipper.bTreatAsArt4 or zipper.treatASJSTemplate then
foldername=foldername
else
foldername=foldername & "\" & replace(Request("zip"),"." & GetFileExtension(Request("zip")),"",1,-1,1)
end if
Response.Redirect ("bs_CreateTemplate.asp?path="& foldername)
end if
else
bContinue=false
end if
set zipper=nothing
if not convertBool(Request.Form ("bRemove")) then
bRemove=false
end if
end if
if not bContinue then%><p align="center"><font color=Red><b>The zip file does not look like a zipped template.</b></font><br /><br /><a href="bs_uploadzip.asp">Try again</a></p><%else%><FORM method="post" ACTION="bs_unzip.asp" id=form1 name=form1> 
<%=QS_secCodeHidden%><input type="hidden" name="zip" value="<%=sanitize(Request("zip"))%>" /><input type="hidden" name="postBack" value="<%=true%>" /><table align=center cellpadding=5 border=0><tr><td class="QSlabel">Name zipfile:</td><td><%=Request("zip")%></td></tr><%if zipper.supportZipper then%><tr><td class="QSlabel">Store in folder:*</td><td><select name="sPath"><%=fe.SelectBoxFolders(server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles")),defaultFolder)%></select></td></tr><tr><td class="QSlabel" valign="top">Template folder name:</td><td><b><%=replace(Request("zip"),"." & GetFileExtension(Request("zip")),"",1,-1,1)%></b></td></tr><%end if%><tr><td class="QSlabel">Remove zip file?</td><td><input type="radio" <%if bRemove then Response.Write " checked "%> value="<%=true%>" name="bRemove" /> Yes 
<input <%if not bRemove then Response.Write " checked "%> type="radio" value="<%=false%>" name="bRemove" /> No</td></tr><tr><td class="QSlabel">&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td>&nbsp;</td><td><INPUT class="art-button" TYPE=SUBMIT VALUE="Continue" id=SUBMIT1 name=SUBMIT1 /></td></tr></table></form><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
