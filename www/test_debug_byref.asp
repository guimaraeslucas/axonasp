<%
Response.Write "Testing ByRef Parameter Passing<br/><hr/>"

Sub ModifyByRef(ByRef value)
    Response.Write "Inside ModifyByRef: value before = " & value & "<br/>"
    value = value * 2
    Response.Write "Inside ModifyByRef: value after = " & value & "<br/>"
End Sub

Sub ModifyByVal(ByVal value)
    Response.Write "Inside ModifyByVal: value before = " & value & "<br/>"
    value = value * 2
    Response.Write "Inside ModifyByVal: value after = " & value & "<br/>"
End Sub

Dim byRefVar
byRefVar = 5
Response.Write "Before ModifyByRef: byRefVar = " & byRefVar & "<br/>"
ModifyByRef(byRefVar)
Response.Write "After ModifyByRef: byRefVar = " & byRefVar & "<br/>"
Response.Write "Expected: 10, Got: " & byRefVar & "<br/><br/>"

Dim byValVar
byValVar = 5
Response.Write "Before ModifyByVal: byValVar = " & byValVar & "<br/>"
ModifyByVal(byValVar)
Response.Write "After ModifyByVal: byValVar = " & byValVar & "<br/>"
Response.Write "Expected: 5, Got: " & byValVar & "<br/>"
%>
