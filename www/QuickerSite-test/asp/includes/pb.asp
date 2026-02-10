
<%class PrivateBot
public filter,folder,only,selectedfiles
private fso, emailDict, file, conn, adox, table, rs, el, arrContent, field, iFilterKey, addEmail, var,textfile,p_folder,errors,ccounter,totalMB
private sub class_initialize
folder="files"
set p_folder=nothing
ccounter=0
totalMB=0
set fso=server.createobject("scripting.filesystemobject")
set emailDict=server.CreateObject ("scripting.dictionary")
set errors=server.CreateObject ("scripting.dictionary")
end sub
public function getFolder
if p_folder is nothing then
set p_folder=fso.getfolder(server.MapPath (folder))
end if
set getFolder=p_folder
end function
public sub deleteFile(filename)
for each file in getFolder.files
if file.name=filename then
file.delete
end if
next
end sub
public function getEmailList()
emailDict.RemoveAll 
'initialize
filter=split(filter,",")
only=split(only,",")
for each file in getFolder.files
On Error Resume Next
err.Clear 
if instr(selectedfiles,file.name)<>0 then
select case lcase(right(file,4))
case ".mdb",".xls","xlsx"
Set conn = Server.Createobject("ADODB.Connection")
select case lcase(right(file,4))
case ".xls","xlsx"
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&file&";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
case ".mdb"
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & file
end select 
if err.number <>0 then
errors.Add file.name,err.Description 
else
totalMB=totalMB+file.size
end if
Set adox = Server.CreateObject("ADOX.Catalog")  
adox.activeConnection = conn 
for each table in adox.tables  
    if table.type="TABLE" then  	    
    
Set rs = server.CreateObject ("adodb.recordset")
rs.CursorType = 1
rs.LockType = 3
rs.ActiveConnection = conn
rs.Open "select * from ["& table.name & "]"
while not rs.EOF 
for each field in rs.Fields
var=rs(field.name)
if not isNull(var) then
if instr(var,"@")<>0 then
arrContent=split(var," ")
checkEmail(arrContent)
end if
end if
next
rs.MoveNext 
wend
Set rs = nothing
        
    end if 
    
next
conn.Close
set conn=nothing
set adox=nothing
case ".txt",".htm","html"
'open as textstream
set textfile=fso.OpenTextFile (file,1)
arrContent=split(cleanup(textfile.readAll)," ")
checkEmail(arrContent)
set textfile=nothing
end select
end if
On Error Goto 0
next
export()
end function
public function export
if errors.Count> 0 then
Response.Write "<ul>"
dim error
for each error in errors
Response.Write "<li><font style='color:Red'>Error: <b>" & error & "</b>: " & errors(error) & "</font></li>"
next
Response.Write "</ul>"
end if
Response.Write "<u>Results</u>:<ul>"
Response.Write "<li><b>" & emailDict.Count & "</b> different email addresses found!</li>"
Response.Write "<li>total nmbr of email addresses found: <b>"&ccounter&"</b></li>"
if totalMB>0 then
Response.Write "<li>total nmbr of MB searched: <b>" & round(totalMB/1024/1024,1) & "</b> MB</li>"
end if
Response.Write "</ul>"
if emailDict.Count>0 then
dim emailNumber, runner
runner=0
emailNumber=emailDict.Count-1
if emailNumber=0 then emailNumber=1
dim strEMail, arrEmail()
redim arrEmail(emailNumber)
for each strEMail in emailDict
arrEmail(runner)=strEMail
runner=runner+1
next
Call QuickSort(arrEmail,0,emailNumber)
Call PrintArray(arrEmail,0,emailNumber)
end if
end function
public function quickSearch(requestVar)
emailDict.RemoveAll 
requestVar=cleanup(requestVar)
arrContent=split(requestVar," ")
checkEmail(arrContent)
export()
end function
private function cleanup(txtpart)
if len(txtpart)<6 then
cleanup=""
exit function
end if
cleanup=txtpart
while right(cleanup,1)="."
cleanup=left(cleanup,len(cleanup)-1)
wend
while instr(cleanup,"..")<>0
cleanup=replace(cleanup,".."," ",1,-1,1)
wend
cleanup=replace(cleanup,"<"," ",1,-1,1)
cleanup=replace(cleanup,">"," ",1,-1,1)
cleanup=replace(cleanup,";"," ",1,-1,1)
cleanup=replace(cleanup,","," ",1,-1,1)
cleanup=replace(cleanup,"'"," ",1,-1,1)
cleanup=replace(cleanup,"/"," ",1,-1,1)
cleanup=replace(cleanup,":"," ",1,-1,1)
cleanup=replace(cleanup,")"," ",1,-1,1)
cleanup=replace(cleanup,"("," ",1,-1,1)
cleanup=replace(cleanup,"]"," ",1,-1,1)
cleanup=replace(cleanup,"["," ",1,-1,1)
cleanup=replace(cleanup,"{"," ",1,-1,1)
cleanup=replace(cleanup,"}"," ",1,-1,1)
cleanup=replace(cleanup,"?"," ",1,-1,1)
cleanup=replace(cleanup,"!"," ",1,-1,1)
cleanup=replace(cleanup,"|"," ",1,-1,1)
cleanup=replace(cleanup,"&"," ",1,-1,1)
cleanup=replace(cleanup,"%"," ",1,-1,1)
cleanup=replace(cleanup,"+"," ",1,-1,1)
cleanup=replace(cleanup,"="," ",1,-1,1)
cleanup=replace(cleanup,"$"," ",1,-1,1)
cleanup=replace(cleanup,". "," ",1,-1,1)
cleanup=replace(cleanup,""""," ",1,-1,1)
cleanup=replace(cleanup,vbcrlf," ",1,-1,1)
cleanup=replace(cleanup,vbcr," ",1,-1,1)
cleanup=replace(cleanup,vblf," ",1,-1,1)
cleanup=replace(cleanup,vbtab," ",1,-1,1)
cleanup=replace(cleanup,vbNewLine," ",1,-1,1)
cleanup=lcase(trim(cleanup))
if len(cleanup)<6 then cleanup=""
end function
' Action: checks if an email is correct. 
' Parameter: Email address 
' Returned value: on success it returns True, else False. 
private Function CheckEmailSyntax(strEmail) 
 
On Error Resume Next
    Dim strArray 
    Dim strItem 
    Dim i 
    Dim c 
  
  
    ' assume the email address is correct  
    CheckEmailSyntax = True 
    
    ' split the email address in two parts: name@domain.ext 
    strArray = Split(strEmail, "@") 
  
    ' if there are more or less than two parts  
    If UBound(strArray) <> 1 Then 
        CheckEmailSyntax = False               
        Exit Function 
    End If 
  
    ' check each part 
    For Each strItem In strArray 
        ' no part can be void 
        If Len(strItem) <= 0 Then 
            CheckEmailSyntax = False           
            Exit Function 
        End If 
        
        ' check each character of the part 
        ' only following "abcdefghijklmnopqrstuvwxyz_-." 
        ' characters and the ten digits are allowed 
        For i = 1 To Len(strItem) 
               c = LCase(Mid(strItem, i, 1)) 
               ' if there is an illegal character in the part 
               If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then 
                   CheckEmailSyntax = False                                    
                   Exit Function 
               End If 
        Next 
   
      ' the first and the last character in the part cannot be . (dot) 
        If Left(strItem, 1) = "." Or Right(strItem, 1) = "." Then 
           CheckEmailSyntax = False	   
           Exit Function 
        End If 
    Next 
  
    ' the second part (domain.ext) must contain a . (dot) 
    If InStr(strArray(1), ".") <= 0 Then 
       CheckEmailSyntax = False	 
        Exit Function 
    End If 
  
    ' check the length oh the extension  
    i = Len(strArray(1)) - InStrRev(strArray(1), ".") 
    ' the length of the extension can be only 2, 3, or 4 
    ' to cover the new "info" extension 
    If i <> 2 And i <> 3 And i <> 4 Then 
        CheckEmailSyntax = False
        Exit Function 
    End If 
    ' after . (dot) cannot follow a . (dot) 
    If InStr(strEmail, "..") > 0 Then 
        CheckEmailSyntax = False
        Exit Function 
    End If 
On Error Goto 0
End Function
private function checkEmail(byval arrContent2)
dim txt
for eL=lbound(arrContent2) to ubound(arrContent2)
txt=arrContent2(eL)
if instr(txt,"@")<>0 then
txt=cleanup(txt)
'check if valid email address
if CheckEmailSyntax(txt) then
ccounter=ccounter+1
'add if not yet added
if not emailDict.Exists (txt) then 
addEmail=true
'remove all the filtered
if isArray(filter) then
for iFilterKey=lbound(filter) to ubound(filter)
if instr(txt,filter(iFilterKey))<>0 then
addEmail=false
exit for
end if
next
end if
'keep the "only"
if isArray(only) then
if addEmail then
for iFilterKey=lbound(only) to ubound(only)
addEmail=false
if instr(txt,only(iFilterKey))<>0 then
addEmail=true
exit for
end if
next
end if
end if
if addEmail then emailDict.Add txt,txt
end if
end if
end if
next
end function
private Sub QuickSort(vec,loBound,hiBound)
  Dim pivot,loSwap,hiSwap,temp
  '== This procedure is adapted from the algorithm given in:
  '==    Data Abstractions & Structures using C++ by
  '==    Mark Headington and David Riley, pg. 586
  '== Quicksort is the fastest array sorting routine for
  '== unordered arrays.  Its big O is  n log n
  '== Two items to sort
  if hiBound - loBound = 1 then
    if vec(loBound) > vec(hiBound) then
      temp=vec(loBound)
      vec(loBound) = vec(hiBound)
      vec(hiBound) = temp
    End If
  End If
  '== Three or more items to sort
  pivot = vec(int((loBound + hiBound) / 2))
  vec(int((loBound + hiBound) / 2)) = vec(loBound)
  vec(loBound) = pivot
  loSwap = loBound + 1
  hiSwap = hiBound
  
  do
    '== Find the right loSwap
    while loSwap < hiSwap and vec(loSwap) <= pivot
      loSwap = loSwap + 1
    wend
    '== Find the right hiSwap
    while vec(hiSwap) > pivot
      hiSwap = hiSwap - 1
    wend
    '== Swap values if loSwap is less then hiSwap
    if loSwap < hiSwap then
      temp = vec(loSwap)
      vec(loSwap) = vec(hiSwap)
      vec(hiSwap) = temp
    End If
  loop while loSwap < hiSwap
  
  vec(loBound) = vec(hiSwap)
  vec(hiSwap) = pivot
  
  '== Recursively call function .. the beauty of Quicksort
    '== 2 or more items in first section
    if loBound < (hiSwap - 1) then Call QuickSort(vec,loBound,hiSwap-1)
    '== 2 or more items in second section
    if hiSwap + 1 < hibound then Call QuickSort(vec,hiSwap+1,hiBound)
End Sub  'QuickSort
private Sub PrintArray(vec,lo,hi)
  
'== Simply print out an array from the lo bound to the hi bound.
Response.Write "<textarea onclick=""javascript:this.select();"" cols=""90"" rows=""15"" id=""pairs"" name=""pairs"">"
Dim i
For i = lo to hi
  Response.Write vec(i) & vbcrlf
Next
Response.Write "</textarea>"
response.write "<script type='text/javascript'>alert('" & hi+1 & " different emails were found.\n\nClick IMPORT now!')</script>"
End Sub  'PrintArray
end class%>
