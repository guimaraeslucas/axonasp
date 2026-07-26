<%
' ============================================================
' lib/music.asp - Music directory resolution and scanning
' ============================================================

' GetMusicDir resolves the music directory from path_aliases.dat.
' Falls back to the default /music/ web folder if no alias is configured.
Function GetMusicDir()
    Dim fso, aliasPath, musicDefault
    musicDefault = Server.MapPath("/music")
    GetMusicDir = musicDefault

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    aliasPath = Server.MapPath("/data") & "\path_aliases.dat"

    If fso.FileExists(aliasPath) Then
        Dim f, line, parts
        Set f = fso.OpenTextFile(aliasPath, 1)
        Do While Not f.AtEndOfStream
            line = Trim(f.ReadLine)
            If Len(line) > 0 And Left(line, 1) <> ";" And Left(line, 1) <> "#" Then
                parts = Split(line, "|")
                If UBound(parts) >= 1 Then
                    If LCase(Trim(parts(0))) = "/music/" Then
                        Dim resolved
                        resolved = Trim(parts(1))
                        If Len(resolved) > 0 And fso.FolderExists(resolved) Then
                            GetMusicDir = resolved
                        End If
                        Exit Do
                    End If
                End If
            End If
        Loop
        f.Close
        Set f = Nothing
    End If

    Set fso = Nothing
End Function

' ScanMusicFolder counts audio files and accumulates total size.
' Results are returned via ByRef parameters.
Sub ScanMusicFolder(ByVal dir, ByRef audioCount, ByRef totalSize)
    Dim fso, folder, file
    audioCount = 0
    totalSize = 0

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(dir) Then
        Set folder = fso.GetFolder(dir)
        For Each file In folder.Files
            If IsAudioFile(fso, file.Name) Then
                audioCount = audioCount + 1
                totalSize = totalSize + file.Size
            End If
        Next
        Set folder = Nothing
    End If
    Set fso = Nothing
End Sub
%>
