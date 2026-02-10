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
dim postBack,fso, folder, files, file, path, bConsiderArtisteer,foldername
postBack=convertBool(Request.Form("postBack"))
bConsiderArtisteer=false
path=Request("path")
set fso=server.CreateObject ("scripting.filesystemobject")
if fso.FolderExists (path) then
set folder=fso.GetFolder (path)
set files=folder.files
if postBack then
	
	if convertBool(Request.form ("bConsiderArtisteer")) then
		bConsiderArtisteer=true
	else
		bConsiderArtisteer=false
	end if
	
	foldername=Request.Form ("foldername")
	
	if Request.Form("file")="" then
		message.AddError("err_mandatory")
	elseif Request.Form("foldername")="" then
		message.AddError("err_mandatory")
	end if

	if not message.haserrors then
	
		dim template
		set template=new cls_template
		template.sName=foldername
		dim textfile
		textfile=fso.OpenTextFile(Request.Form("file"),1).ReadAll 
		template.sValue=textfile
		folder=replace(path,server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles")),"",1,-1,1) & "/"
		folder=replace(folder,"\","/",1,-1,1)
		folder=C_VIRT_DIR & Application("QS_CMS_userfiles") & folder
		folder=replace(folder,"//","/",1,-1,1)
		
		if bConsiderArtisteer then
			template.bSetContactVar=true
			template.bSetFooterVar=true
			template.bSetCustomHL=true
			template.bSetRSSLink=true
			template.treatAsArtisteer(folder)
		end if
		
		template.initWAP
		template.initPrint
		template.initMobile
		template.initEmail
		template.initPrint
		
		if template.save() then
			Response.Redirect ("bs_templateList.asp?fbMessage=fb_templateadded")
		else
			message.AddError("err_mandatory")
		end if
	end if
	
else

	foldername=folder.name	
	
end if

if files.count=0 then
fso.Deletefolder (path)%><p align="center">Your zip file has no files in its root. You have to re-zip and make sure the root of the zip file contains the html of the template.<p><%else%><FORM method="post" ACTION="bs_CreateTemplate.asp" id=form1 name=form1> 
<%=QS_secCodeHidden%><input type="hidden" name="path" value="<%=sanitize(Request("path"))%>" /><input type="hidden" name="postBack" value="<%=true%>" /><table align=center cellpadding=5 border=0><tr><td class="QSlabel">Template name:*</td><td><input type=text name="foldername" value="<%=sanitize(foldername)%>" size="30" /></td></tr><tr><td valign="top" class="QSlabel">Select the html template:*</td><td valign="top">
<%dim checked,continue,readAll,showFile
continue=true
for each file in files
	showFile=false
	if not postback then
		if lcase(file.name)="page.html" or lcase(file.name)="index.html" then
			bConsiderArtisteer=true
			showFile=true
		end if
		if not showFile then
			if instr(file.name,".htm")<>0 then
				bConsiderArtisteer=true
				showFile=true
			end if
		end if
		if lcase(file.name)="blog.html" or lcase(file.name)="home.html" then
			showfile=false
		end if
		if not allowedFileTypes.exists(lcase(convertStr(GetFileExtension(file)))) then
			on error resume next
			fso.DeleteFile (file)
			on error goto 0			
		end if
		if continue then
			select case lcase(convertStr(file.name))
				case "page.html","page.htm","index.html","index.htm"
					checked=" checked "
					continue=false
				case else
					checked=""
			end select
		else
			checked=""
		end if
	end if
	if showFile then
	%>
		<input type="radio" <%=checked%> <%if sanitize(file)=Request.Form ("file") then Response.Write " checked "%> value="<%=sanitize(file)%>" name="file" /><%=file.name%><br />
	<%
	end if
next

%></td></tr><tr><td class="QSlabel">Treat as a JStemplate template?</td><td><input name="bConsiderArtisteer" type="checkbox" value="<%=true%>" <%if bConsiderArtisteer then Response.Write " checked "%> /></td></tr><tr><td class="QSlabel">&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td>&nbsp;</td><td><INPUT class="art-button" TYPE=SUBMIT VALUE="Continue" id=SUBMIT1 name=SUBMIT1 /></td></tr></table></form><%end if
else%><p align="center">The folder cannot be found. Create a new zip file, and avoid using special characters in its name.</p><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
