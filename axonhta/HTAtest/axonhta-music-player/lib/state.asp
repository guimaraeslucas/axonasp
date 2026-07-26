<%
' ============================================================
' lib/state.asp - Playback state and path alias persistence
' ============================================================

' LoadState reads saved playback state from data/state.dat.
' Returns idx and vol via ByRef parameters with sensible defaults.
Sub LoadState(ByRef idx, ByRef vol)
    idx = 0
    vol = 70

    Dim fso, statePath
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    statePath = Server.MapPath("/data") & "\state.dat"

    If fso.FileExists(statePath) Then
        Dim f, line, parts
        Set f = fso.OpenTextFile(statePath, 1)
        If Not f.AtEndOfStream Then
            line = Trim(f.ReadLine)
            parts = Split(line, "|")
            If UBound(parts) >= 1 Then
                If IsNumeric(parts(0)) Then idx = CLng(parts(0))
                If IsNumeric(parts(1)) Then vol = CLng(parts(1))
            End If
        End If
        f.Close
        Set f = Nothing
    End If

    Set fso = Nothing
End Sub

' SaveState writes playback state to data/state.dat.
Sub SaveState(ByVal idx, ByVal vol)
    Dim fso, dataDir, statePath
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    dataDir = Server.MapPath("/data")
    If Not fso.FolderExists(dataDir) Then fso.CreateFolder dataDir
    statePath = dataDir & "\state.dat"

    Dim f
    Set f = fso.CreateTextFile(statePath, True)
    f.WriteLine idx & "|" & vol
    f.Close
    Set f = Nothing
    Set fso = Nothing
End Sub

' SaveMusicDir writes or replaces the /music/ alias in path_aliases.dat.
Sub SaveMusicDir(ByVal dir)
    Dim fso, dataDir, aliasPath
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    dataDir = Server.MapPath("/data")
    If Not fso.FolderExists(dataDir) Then fso.CreateFolder dataDir
    aliasPath = dataDir & "\path_aliases.dat"

    Dim existingLines, line, linePrefix
    existingLines = ""

    If fso.FileExists(aliasPath) Then
        Dim f
        Set f = fso.OpenTextFile(aliasPath, 1)
        Do While Not f.AtEndOfStream
            line = Trim(f.ReadLine)
            If Len(line) > 0 And Left(line, 1) <> ";" And Left(line, 1) <> "#" Then
                linePrefix = Trim(Split(line, "|")(0))
                If LCase(linePrefix) <> "/music/" Then
                    existingLines = existingLines & line & vbNewLine
                End If
            ElseIf Len(line) > 0 Then
                existingLines = existingLines & line & vbNewLine
            End If
        Loop
        f.Close
        Set f = Nothing
    End If

    existingLines = existingLines & "/music/|" & Trim(dir) & vbNewLine

    Dim out
    Set out = fso.CreateTextFile(aliasPath, True)
    out.Write existingLines
    out.Close
    Set out = Nothing
    Set fso = Nothing
End Sub
%>
