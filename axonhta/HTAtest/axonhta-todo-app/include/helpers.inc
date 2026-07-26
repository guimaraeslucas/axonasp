<%
' ============================================================
' helpers.inc - Task CRUD operations and utilities
' Data format: id|title|notes|priority|status|created
'   priority: low / medium / high
'   status:   active / done
' ============================================================

' --- Path to the data file
Function TaskFile()
    TaskFile = Server.MapPath("/data") & "\tasks.dat"
End Function

' --- HTML-encode a value
Function Esc(ByVal s)
    Esc = Server.HTMLEncode(CStr(s))
End Function

' --- Return CSS class for priority badge
Function PriorityClass(ByVal p)
    Select Case LCase(Trim(p))
        Case "high":   PriorityClass = "badge-high"
        Case "medium": PriorityClass = "badge-medium"
        Case "low":    PriorityClass = "badge-low"
        Case Else:     PriorityClass = "badge-medium"
    End Select
End Function

' --- Format a timestamp
Function Timestamp()
    Timestamp = Year(Now) & "-" & _
        Right("0" & Month(Now), 2) & "-" & _
        Right("0" & Day(Now), 2) & " " & _
        Right("0" & Hour(Now), 2) & ":" & _
        Right("0" & Minute(Now), 2)
End Function

' --- Read entire file; returns "" if not found
Function ReadFile(ByVal path)
    On Error Resume Next
    Dim fso, f
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then
        ReadFile = ""
        Set fso = Nothing
        Exit Function
    End If
    Set f = fso.OpenTextFile(path, 1)
    ReadFile = f.ReadAll
    f.Close
    Set f = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Function

' --- Write entire content to file (overwrite)
Sub WriteFile(ByVal path, ByVal content)
    On Error Resume Next
    Dim fso, f
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(path, True)
    f.Write content
    f.Close
    Set f = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Sub

' ============================================================
' CRUD: Create a new task
' Returns the new task ID
' ============================================================
Function CreateTask(ByVal title, ByVal notes, ByVal priority)
    Dim fso, f, nextId, raw, lines, arrLines, maxId, i
    Dim cleanTitle, cleanNotes

    ' Sanitize pipe characters from input
    cleanTitle  = Replace(Replace(title, vbCrLf, " "), "|", "/")
    cleanNotes  = Replace(Replace(notes, vbCrLf, " "), "|", "/")

    ' Read existing data to find max ID
    raw = ReadFile(TaskFile())
    nextId = 1
    If Len(raw) > 0 Then
        arrLines = Split(raw, vbNewLine)
        For i = 0 To UBound(arrLines)
            If Len(Trim(arrLines(i))) > 0 Then
                Dim parts
                parts = Split(arrLines(i), "|")
                If UBound(parts) >= 0 Then
                    If IsNumeric(parts(0)) Then
                        If CLng(parts(0)) >= nextId Then
                            nextId = CLng(parts(0)) + 1
                        End If
                    End If
                End If
            End If
        Next
    End If

    ' Append new line
    Dim newLine
    newLine = nextId & "|" & cleanTitle & "|" & cleanNotes & "|" & priority & "|active|" & Timestamp()

    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(TaskFile(), 8, True)
    f.WriteLine newLine
    f.Close
    Set f = Nothing
    Set fso = Nothing
    On Error GoTo 0

    CreateTask = nextId
End Function

' ============================================================
' CRUD: Update a task by ID
' ============================================================
Sub UpdateTask(ByVal taskId, ByVal title, ByVal notes, ByVal priority, ByVal status)
    Dim raw, arrLines, result, i, parts, cleanTitle, cleanNotes

    cleanTitle = Replace(Replace(title, vbCrLf, " "), "|", "/")
    cleanNotes = Replace(Replace(notes, vbCrLf, " "), "|", "/")

    raw = ReadFile(TaskFile())
    If Len(raw) = 0 Then Exit Sub

    arrLines = Split(raw, vbNewLine)
    result = ""
    For i = 0 To UBound(arrLines)
        If Len(Trim(arrLines(i))) > 0 Then
            parts = Split(arrLines(i), "|")
            If UBound(parts) >= 0 And IsNumeric(parts(0)) Then
                If CLng(parts(0)) = CLng(taskId) Then
                    ' Preserve original creation timestamp
                    Dim created
                    created = ""
                    If UBound(parts) >= 5 Then created = parts(5)
                    arrLines(i) = taskId & "|" & cleanTitle & "|" & cleanNotes & "|" & priority & "|" & status & "|" & created
                End If
                result = result & arrLines(i) & vbNewLine
            End If
        End If
    Next

    WriteFile TaskFile(), result
End Sub

' ============================================================
' CRUD: Toggle task status between active and done
' ============================================================
Sub ToggleTask(ByVal taskId)
    Dim raw, arrLines, result, i, parts
    raw = ReadFile(TaskFile())
    If Len(raw) = 0 Then Exit Sub

    arrLines = Split(raw, vbNewLine)
    result = ""
    For i = 0 To UBound(arrLines)
        If Len(Trim(arrLines(i))) > 0 Then
            parts = Split(arrLines(i), "|")
            If UBound(parts) >= 4 And IsNumeric(parts(0)) Then
                If CLng(parts(0)) = CLng(taskId) Then
                    If parts(4) = "done" Then
                        parts(4) = "active"
                    Else
                        parts(4) = "done"
                    End If
                    arrLines(i) = Join(parts, "|")
                End If
                result = result & arrLines(i) & vbNewLine
            End If
        End If
    Next

    WriteFile TaskFile(), result
End Sub

' ============================================================
' CRUD: Delete a task by ID
' ============================================================
Sub DeleteTask(ByVal taskId)
    Dim raw, arrLines, result, i, parts
    raw = ReadFile(TaskFile())
    If Len(raw) = 0 Then Exit Sub

    arrLines = Split(raw, vbNewLine)
    result = ""
    For i = 0 To UBound(arrLines)
        If Len(Trim(arrLines(i))) > 0 Then
            parts = Split(arrLines(i), "|")
            If UBound(parts) >= 0 And IsNumeric(parts(0)) Then
                If CLng(parts(0)) <> CLng(taskId) Then
                    result = result & arrLines(i) & vbNewLine
                End If
            End If
        End If
    Next

    WriteFile TaskFile(), result
End Sub

' ============================================================
' Read: Get a single task by ID; returns array or Empty
' ============================================================
Function GetTask(ByVal taskId)
    Dim raw, arrLines, i, parts
    raw = ReadFile(TaskFile())
    If Len(raw) = 0 Then
        GetTask = Empty
        Exit Function
    End If

    arrLines = Split(raw, vbNewLine)
    For i = 0 To UBound(arrLines)
        If Len(Trim(arrLines(i))) > 0 Then
            parts = Split(arrLines(i), "|")
            If UBound(parts) >= 5 And IsNumeric(parts(0)) Then
                If CLng(parts(0)) = CLng(taskId) Then
                    GetTask = parts
                    Exit Function
                End If
            End If
        End If
    Next

    GetTask = Empty
End Function

' ============================================================
' Read: Get all tasks as an array of arrays
' ============================================================
Function GetAllTasks()
    Dim raw, arrLines, i, parts, result(), count
    raw = ReadFile(TaskFile())
    count = 0
    If Len(raw) = 0 Then
        GetAllTasks = Array()
        Exit Function
    End If

    arrLines = Split(raw, vbNewLine)
    ' First pass: count valid lines
    For i = 0 To UBound(arrLines)
        If Len(Trim(arrLines(i))) > 0 Then
            parts = Split(arrLines(i), "|")
            If UBound(parts) >= 5 Then
                count = count + 1
            End If
        End If
    Next

    If count = 0 Then
        GetAllTasks = Array()
        Exit Function
    End If

    ' Second pass: collect valid tasks
    ReDim result(count - 1)
    Dim idx : idx = 0
    For i = 0 To UBound(arrLines)
        If Len(Trim(arrLines(i))) > 0 Then
            parts = Split(arrLines(i), "|")
            If UBound(parts) >= 5 Then
                Set result(idx) = parts
                idx = idx + 1
            End If
        End If
    Next

    GetAllTasks = result
End Function

' --- Count tasks by status
Function CountByStatus(ByVal allTasks, ByVal status)
    Dim i, count
    count = 0
    If IsArray(allTasks) Then
        For i = 0 To UBound(allTasks)
            If allTasks(i)(4) = status Then count = count + 1
        Next
    End If
    CountByStatus = count
End Function
%>