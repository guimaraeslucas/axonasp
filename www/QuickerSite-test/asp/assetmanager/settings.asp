<%
Dim arrBaseFolder
Redim arrBaseFolder(4)
Dim arrBaseName
Redim arrBaseName(4)

Dim bReturnAbsolute
bReturnAbsolute=false

arrBaseFolder(0)=Application("QS_CMS_C_VIRT_DIR") & Application("QS_CMS_userfiles") & session("userfolderpath") 'Use "Relative to Root" Path
arrBaseName(0)="Userfiles"

arrBaseFolder(1)=""
arrBaseName(1)=""

arrBaseFolder(2)=""
arrBaseName(2)=""

arrBaseFolder(3)=""
arrBaseName(3)=""
%>