<%
Public Function GetDynamicRS
Set GetDynamicRS = server.CreateObject ("adodb.recordset")
GetDynamicRS.CursorType = 1				
GetDynamicRS.LockType = 3
set GetDynamicRS.ActiveConnection = getConn()		

End Function

Public function getConn()	

	Set getConn = Server.Createobject("ADODB.Connection")	
	getConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath(C_DIRECTORY_QUICKERSITE & "/asp/iismanager/iis.resx")												

end function

FUNCTION State2Desc( nState )
    SELECT CASE nState
    CASE 1
        State2Desc = "Starting"
    CASE 2
        State2Desc = "Started"
    CASE 3
        State2Desc = "Stopping"
    CASE 4
        State2Desc = "Stopped"
    CASE 5
        State2Desc = "Pausing"
    CASE 6
        State2Desc = "Paused"
    CASE 7
        State2Desc = "Continuing"
    CASE ELSE
        State2Desc = "Unknown state"
    END SELECT

END FUNCTION


sub loadIIS()

	on error resume next
	
	dim IISOBJ,siteObject,objWeb,cont,k,m,i,bs,vdn,vdp,obj
	Set IISOBJ = getObject("IIS://LOCALHOST/W3SVC")
	
	if err.number<>0 then exit sub

	getConn.execute("delete from sites")

	dim rs
	set rs=GetDynamicRS
	rs.open "select * from sites"

	For each siteObject in IISOBJ

		if (siteObject.Class = "IIsWebServer") then

			rs.addNew()
			rs("name")=siteObject.name
			rs("ServerComment")=siteObject.ServerComment
			rs("ServerState")=State2Desc(siteObject.ServerState)

			set objWeb = getObject(siteObject.adsPath & "/ROOT")

			'cont=true
			'for k=lbound(siteObject.httpErrors) to ubound(siteObject.httpErrors)
			
			'	for m=0 to 50
			'		if instr(siteObject.httpErrors(k)(m),"default.asp")<>0 then
						'response.write "<li>" & siteObject.httpErrors(k)(m) & "</li>"
			'			cont=false
			'			exit for
			'		end if
			'	next
				
			'	if not cont then exit for
				
			'next

			if objWeb.KeyType="" then
				objWeb.keytype="IIsWebVirtualDir"
				objWeb.SetInfo()
			end if

			rs("Path")=objWeb.Path
			
			for i=lbound(siteObject.serverbindings) to ubound(siteObject.serverbindings)
				bs=bs& siteObject.serverbindings(i)(i) & vbcrlf
			next

			rs("serverbindings")=bs
			bs=""

			for each obj in objWeb
				if obj.class="IIsWebVirtualDir" then
					vdn=vdn&obj.name &vbcrlf
					vdp=vdp&obj.path &vbcrlf
				end if
			next

			rs("QSVD")=vdn
			rs("QSPath")=vdp
			vdn=""
			vdp=""

			rs.update()

			set objWeb=nothing
			
		end if	

	next
	
	err.clear

	on error goto 0

	set rs=nothing

end sub
%>