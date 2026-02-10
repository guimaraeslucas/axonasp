
<%Class cls_LogonEdit
Public authenticatedAsAdmin
Public authenticatedIntranet
Public mode
Private p_contact
Private p_UserIP
Private Sub Class_Initialize
Set p_contact = New cls_contact
p_contact.sEmail=left(request.cookies("sEmail"),50)
authenticatedIntranet=false
logonAdmin request.querystring("apw")
logonIntranet request.cookies("sEmail"), decrypt(request.cookies("sPw"))
if not authenticatedIntranet then
if Request.QueryString ("sEmail")<>"" and Request.QueryString ("sPw")<>"" then
mode="mode1"
logonIntranet left(Request.QueryString ("sEmail"),100), left(Request.QueryString ("sPw"),100)
end if
end if
p_UserIP=UserIP
End Sub
Public Function logon(byRef password)
	logon=false
	if password<>"" then
		if (password=customer.adminPassword and not isLeeg(customer.adminPassword)) or (password=customer.secondAdmin.sPassword and not isLeeg(customer.secondAdmin.sPassword)) then
			
			Session(cId & "isAUTHENTICATED")	= true
			
			if QS_enableCookieMode then 
				response.cookies.item(cId & "hfsdsiiqqssdfjf")	= password
			end if
			
			if password=customer.secondAdmin.sPassword then
				Session(cId & "isAUTHENTICATEDSecondAdmin")= true
			end if
			
			Application("QS_CMS_C_VIRT_DIR")	= C_VIRT_DIR
			Application("QS_CMS_C_DIRECTORY_QUICKERSITE")	= C_DIRECTORY_QUICKERSITE
			Application("QS_ASPX")	= QS_ASPX
			logon=true
			application("bsLoginCount"&UserIP)=0
		End If
	End if
End Function
public property get currentPW
if Session(cId & "isAUTHENTICATEDSecondAdmin") then
currentPW=customer.secondAdmin.sPassword
exit property
end if
if Session(cId & "isAUTHENTICATED") then
currentPW=customer.adminPassword
exit property
end if
end property
Public Function logonAdmin(byRef password)
	if password<>"" then
		if convertStr(password)=Encrypt(C_ADMINPASSWORD) or password=sha256(C_ADMINPASSWORD) then
			Session(cId & "isAUTHENTICATEDasADMIN")=true
			application("adminLoginCount"&p_UserIP)=0
		else
			lockAdmin()
			message.AddError("err_login")
		End If
	end if
	logonAdmin = convertBool(Session(cId & "isAUTHENTICATEDasADMIN"))
End Function

Public function lockAdmin()
if C_ADMINPASSWORD="" then 
Response.Write "This section is disabled (no admin-password specified in asp/config/web_config.asp)."
application("adminLoginCount"&p_UserIP)=0
Response.End 
end if
if application("adminLoginCount"&p_UserIP)>(QS_number_of_allowed_attempts_to_login-1)  then
On Error Resume Next
Err.Raise 11,"not_available.asp","Someone has tried to get access to the ADMIN area more then "&QS_number_of_allowed_attempts_to_login&" times. His/Her IP is now blocked."
if application("mailSent"&p_UserIP)<>"true" then
application("mailSent"&p_UserIP)="true"
dumpError "Someone is trying to connect to the admin section...", err
end if
On Error Goto 0
Err.Raise 11,"not_available.asp","You have attempted to get access "&QS_number_of_allowed_attempts_to_login&" times. You IP is now blocked."
Response.End 
end if
end function
Public function lockBSAdmin()
if application("bsLoginCount"&p_UserIP)>(QS_number_of_allowed_attempts_to_login-1)  then
On Error Resume Next
Err.Raise 11,"not_available.asp","Someone has tried to get access to the BACKSITE area more then "&QS_number_of_allowed_attempts_to_login&" times. His/Her IP is now blocked."
if application("mailSent"&p_UserIP)<>"true" then
application("mailSent"&p_UserIP)="true"
dumpError "Someone is trying to connect to the backsite section...", err
end if
On Error Goto 0
Err.Raise 11,"not_available.asp","You have attempted to get access "&QS_number_of_allowed_attempts_to_login&" times. You IP is now blocked."
Response.End 
end if
end function
Public Function logonItem(byRef password,byref pageObj)
logonItem=true
password=lcase(convertStr(password))
if not isLeeg(pageObj.sPw)  then
logonItem=false
dim passwords
passwords=split(Request.Cookies("pagePW"),"delimiter")
dim i
for i=lbound(passwords) to ubound (passwords)
if decrypt(passwords(i))=pageObj.sPw then
logonItem = true
exit function
End If
next
if convertStr(password)=convertStr(pageObj.sPw) then
Response.cookies("pagePW") = Request.Cookies("pagePW")& encrypt(password) & "delimiter"
logonItem = true
exit function
end if
end if
End Function
Public Function logonIntranet(byRef sEmail, byRef sPw)
authenticatedIntranet = false
if not isLeeg(sEmail) and not isLeeg(sPw) then 'alleen als er iets ingevuld is!
Dim record
sEmail=convertStr(trim(sEmail))
sPw=convertStr(trim(sPw))
dim sql
select case convertGetal(customer.iLoginMode)
case 0
sql="select iId from tblContact where iCustomerID="& cId &" and sEmail='"& cleanUp(sEmail) & "' and sPw='" & cleanUp(sPw) & "' and iStatus>"&cs_silent
case 1
sql="select iId from tblContact where iCustomerID="& cId &" and sNickName='"& cleanUp(sEmail) & "' and sPw='" & cleanUp(sPw) & "' and iStatus>"&cs_silent
end select
Set record = db.execute(sql)
If not record.eof Then
p_contact.pick(ConvertLng(record.fields("iId")))
'de lastloginTS bewaren!
p_contact.lastLoginTSSave()
select case mode
case "mode1"
response.cookies("sPw") = encrypt(p_contact.sPw)
Response.Cookies("sPw").expires=dateAdd("d",365,date())
response.cookies("sEmail") = p_contact.sEmailorNickname
Response.Cookies("sEmail").expires=dateAdd("d",365,date())
Response.Cookies("mode")="mode1"
Response.Cookies("mode").expires=dateAdd("d",365,date())
case "mode2"
response.cookies("sPw") = encrypt(p_contact.sPw)
response.cookies("sEmail") = p_contact.sEmailorNickname
Response.Cookies("sEmail").expires=dateAdd("d",365,date())
Response.Cookies("mode")="mode2"
Response.Cookies("mode").expires=dateAdd("d",365,date())
case "mode3"
response.cookies("sPw") = encrypt(p_contact.sPw)
response.cookies("sEmail") = p_contact.sEmailorNickname
Response.Cookies("mode")="mode3"
Response.Cookies("mode").expires=dateAdd("d",365,date())
end select
authenticatedIntranet = true
Session(cId & "isAUTHENTICATEDIntranet")= true
End If
Set record = nothing
end if
logonIntranet=authenticatedIntranet
getUFP()
End Function
public function getAllPasswordsCS
getAllPasswordsCS="''"
dim passwords
passwords=split(Request.Cookies("pagePW"),"delimiter")
dim i
for i=lbound(passwords) to ubound (passwords)
getAllPasswordsCS=getAllPasswordsCS & ",'" & left(cleanUpStr(decrypt(passwords(i))),15) & "'"
next
end function
Public Function logoff
response.Cookies("sEmail") = ""
response.Cookies("sPw") = ""
application("adminLoginCount"&p_UserIP)=0
Session(cId & "isAUTHENTICATEDasADMIN")	= false
Session(cId & "isAUTHENTICATEDasUSER")	= false
Session(cId & "isAUTHENTICATEDSecondAdmin")	= false
End Function
Public Function logoffBO
Application(QS_CMS_FCK_allowedIP)=""
Application(QS_CMS_cacheBOMenu & customer.secondAdmin.sPassword)=""
Application(QS_CMS_cacheBOMenu & customer.adminPassword)=""
Session.Abandon ()
End Function
Public Property Get contact
if p_contact is nothing then
set p_contact=new cls_contact
end if
Set contact = p_contact
End Property
Public function resetPW(sEmail)
resetPW=false
if checkEmailSyntax(trim(convertStr(sEmail))) then
dim sql, record
sql="select iId from tblContact where iCustomerID="& cId &" and sEmail='"& cleanUp(sEmail)  &"' and iStatus>"&cs_silent
 
Set record = db.execute(sql)
If not record.eof Then
dim resetPWContact
set resetPWContact=new cls_contact
resetPWContact.pick(convertGetal(record(0)))
resetPWContact.resetPW()
resetPW=true
set resetPWContact=nothing
end if
set record=nothing
end if
end function
public sub hasaccess(bool)
if currentPW=customer.secondAdmin.sPassword then
if not bool then
Response.Redirect ("bs_default.asp")
end if
end if
end sub
public function getUFP 'getUserFilesPath
Application("QS_CMS_C_VIRT_DIR")	= C_VIRT_DIR
Application("QS_CMS_C_DIRECTORY_QUICKERSITE")	= C_DIRECTORY_QUICKERSITE
Application("QS_ASPX")	= QS_ASPX
if Session(cId & "isAUTHENTICATED") then
session("userfolderpath")=""
getUFP=""
elseif authenticatedIntranet then
dim fso2
set fso2=server.CreateObject ("scripting.filesystemobject")
if not fso2.FolderExists (server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles")) then
fso2.CreateFolder (server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles"))
end if
if not fso2.FolderExists (server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/"& contact.iId)) then
fso2.CreateFolder (server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "userfiles/"& contact.iId))
end if
session("userfolderpath")="userfiles/" & contact.iId & "/"
getUFP="userfiles/" & contact.iId & "/"
end if
end function
End Class%>
