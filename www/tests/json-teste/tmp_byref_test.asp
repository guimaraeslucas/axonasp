<%
Option Explicit
Dim i_props, i_count, i_cap

' Simulate class_initialize
ReDim i_props(-1)
i_count = 0
i_cap = 0

' Simulate add x3 (force a resize at count=4)
Dim k
For k = 0 To 2
    If i_count >= i_cap Then
        ReDim Preserve i_props(i_cap * 1.2 + 1)
        i_cap = UBound(i_props) + 1
    End If
    i_props(i_count) = "item" & k
    i_count = i_count + 1
Next

Response.Write "count=" & i_count & " cap=" & i_cap & " ubound=" & UBound(i_props) & "<br>"

' Now simulate ArraySlice on index 1 (ByRef arr not propagating)
' If ByRef doesn't work, i_props won't be ReDim'd back
Dim arr_copy
arr_copy = i_props  ' simulate ByVal pass
i_count = i_count - 1
' Would do: ReDim Preserve arr_copy(i_count * 1.2 + 1) ... but NOT i_props

' Now try access
Dim j
For j = 0 To i_count - 1
    Response.Write j & ": " & i_props(j) & "<br>"
Next

Response.Write "Done"
%>
