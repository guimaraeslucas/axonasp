<%
if request.form("yourname")="" then
	response.write  "Oeps... you forgot to type your name... Try again!"
else
	dim i
	for i = 6 to 1 step -1
		response.write  "<h" & i & ">Hello " & sanitize(URLDecode(request.form("yourname"))) & "!</h" & i & ">"
	next
end if

Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If
	
    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ",1,-1,1)
	
    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")
	
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If
	
    URLDecode = sOutput
End Function

function sanitize(sValue)
	
	if isNull(sValue) then
		sanitize=""
	else
		sanitize=replace(sValue,"""","&quot;",1,-1,1)
		sanitize=replace(sanitize,"<","&lt;",1,-1,1)
		sanitize=replace(sanitize,">","&gt;",1,-1,1)
	end if

end function
%>