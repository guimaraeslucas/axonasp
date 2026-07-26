<!-- #include virtual="/lib/helpers.asp" -->
<!-- #include virtual="/lib/music.asp" -->
<!-- #include virtual="/lib/state.asp" -->
<%
' ============================================================
' api/playlist.asp - HTMX endpoint: returns playlist fragment
' Handles both GET (rescan) and POST (save_dir + rescan)
' ============================================================

' Handle save_dir POST
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Request.Form("action") = "save_dir" Then
        SaveMusicDir Request.Form("dir")
    End If
End If

' Resolve current music directory and scan
Dim musicDir, audioCount, totalSize, savedIdx, savedVol
musicDir = GetMusicDir()
audioCount = 0
totalSize = 0
ScanMusicFolder musicDir, audioCount, totalSize

' Load saved state and clamp index
LoadState savedIdx, savedVol
If savedIdx >= audioCount Then savedIdx = 0
If savedIdx < 0 Then savedIdx = 0
%>
<div id="playlist-area">
  <div class="playlist-header">
    <span>Playlist (<%=audioCount%>)</span>
    <div class="header-actions">
      <button class="btn-add" onclick="openFolderBrowser()" title="Music folder settings">
        <svg viewBox="0 0 24 24" width="16" height="16"><path fill="currentColor" d="M10 4H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2z"/></svg>
        Folder
      </button>
      <a class="btn-add"
         hx-get="/api/playlist.asp"
         hx-target="#playlist-area"
         hx-swap="outerHTML"
         title="Rescan music folder">
        <svg viewBox="0 0 24 24" width="16" height="16"><path fill="currentColor" d="M17.65 6.35A7.958 7.958 0 0 0 12 4c-4.42 0-7.99 3.58-7.99 8s3.57 8 7.99 8c3.73 0 6.84-2.55 7.73-6h-2.08A5.99 5.99 0 0 1 12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6c1.66 0 3.14.69 4.22 1.78L13 11h7V4l-2.35 2.35z"/></svg>
        Rescan
      </a>
    </div>
  </div>

  <% If audioCount > 0 Then %>
  <div class="playlist-info">
    <span><%=audioCount%> files, <%=FmtSize(totalSize)%></span>
    <span>Folder: <%= HtmlEsc(musicDir) %></span>
  </div>
  <% End If %>

  <div class="playlist">
<%
If audioCount > 0 Then
    Dim fso, folder, file, ti
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(musicDir)
    ti = 0
    For Each file In folder.Files
        If IsAudioFile(fso, file.Name) Then
%>
    <div class="playlist-item<% If ti = savedIdx Then %> active<% End If %>"
         data-idx="<%=ti%>"
         data-url="/music/<%= JSEsc(file.Name) %>"
         @click="play($el)">
      <div class="item-index"><%=ti + 1%></div>
      <div class="item-name"><%= HtmlEsc(file.Name) %></div>
    </div>
<%
            ti = ti + 1
        End If
    Next
    Set folder = Nothing
    Set fso = Nothing
Else
%>
    <div class="playlist-empty">
      <svg viewBox="0 0 24 24" width="32" height="32" style="opacity:0.3"><path fill="currentColor" d="M12 3v10.55c-.59-.34-1.27-.55-2-.55C7.79 13 6 14.79 6 17s1.79 4 4 4 4-1.79 4-4V7h4V3h-6z"/></svg>
      <p>No audio files found</p>
      <p class="sub">Click Folder to choose a music directory</p>
    </div>
<%
End If
%>
  </div>
</div>
