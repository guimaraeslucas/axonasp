
<%Function decrypt(byval sIn)
dim x, y, abfrom, abto
decrypt="": ABFrom = ""   
    
For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next 
For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next 
For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next 
abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
For x=1 to Len(sin): y=InStr(abto, Mid(sin, x, 1))
    If y = 0 then
        decrypt = decrypt & Mid(sin, x, 1)
    Else
        decrypt = decrypt & Mid(abfrom, y, 1)
    End If
Next
if instr(decrypt,"gbqmdfk")<>0 and instr(decrypt,"ghbnchom")<>0 then
decrypt=replace(decrypt,"gbqmdfk","")
decrypt=replace(decrypt,"ghbnchom","")   
'elseif len(decrypt)>5 then
'do nothing
elseif isNumeriek(decrypt) then
decrypt=cstr(round(convertGetal(decrypt)/111,0))
end if
   
    
End Function
Function EnCrypt(byval sIn)
'if len(sIn)<6 then
if isNumeriek(sIn) then
sIn=convertGetal(sIn)*111
end if
'end if
sIn=convertStr(sIn)
dim x, y, abfrom, abto
EnCrypt="": ABFrom = ""
For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next 
For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next 
For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next 
abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
For x=1 to Len(sin): y = InStr(abfrom, Mid(sin, x, 1))
    If y = 0 Then
         EnCrypt = EnCrypt & Mid(sin, x, 1)
    Else
         EnCrypt = EnCrypt & Mid(abto, y, 1)
    End If
Next
End Function%>
