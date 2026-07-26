<!-- #include virtual="/lib/helpers.asp" -->
<!-- #include virtual="/lib/music.asp" -->
<%
' ============================================================
' api/browse.asp - HTMX endpoint: folder browser fragment
' Query param: dir = directory to list (default: music dir)
' ============================================================

Dim browseDir, fso, parentDir
browseDir = Request.QueryString("dir")
If browseDir = "" Then browseDir = GetMusicDir()

Set fso = Server.CreateObject("Scripting.FileSystemObject")

' Resolve to absolute path
If fso.FolderExists(browseDir) Then
    browseDir = fso.GetFolder(browseDir).Path
Else
    ' Fallback to server root
    browseDir = Server.MapPath("/")
End If

' Compute parent directory
parentDir = ""
If Len(browseDir) > 3 Then ' Not a drive root like C:\
    parentDir = fso.GetParentFolderName(browseDir)
End If
%>
<div class="fb-header">
  <span class="fb-title">Select Music Folder</span>
  <button class="fb-close" onclick="closeFolderBrowser()">&times;</button>
</div>
<div class="fb-path" title="<%= HtmlEsc(browseDir) %>"><%= HtmlEsc(browseDir) %></div>
<div class="fb-list">
<% If parentDir <> "" And fso.FolderExists(parentDir) Then %>
  <div class="fb-up" hx-get="/api/browse.asp?dir=<%= Server.URLEncode(parentDir) %>"
       hx-target="#folder-browser-body" hx-swap="innerHTML">
    <svg viewBox="0 0 24 24" width="18" height="18"><path fill="currentColor" d="M4 11h12.17l-3.59-3.59L14 6l6 6-6 6-1.41-1.41L16.17 13H4z" transform="rotate(-90 12 12)"/></svg>
    ..
  </div>
<% End If %>
<%
Dim folder, subfolder, hasSubdirs
Set folder = fso.GetFolder(browseDir)
hasSubdirs = False
For Each subfolder In folder.SubFolders
    hasSubdirs = True
%>
  <div class="fb-item" hx-get="/api/browse.asp?dir=<%= Server.URLEncode(subfolder.Path) %>"
       hx-target="#folder-browser-body" hx-swap="innerHTML">
    <svg viewBox="0 0 24 24" width="18" height="18"><path fill="currentColor" d="M10 4H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2z"/></svg>
    <%= HtmlEsc(subfolder.Name) %>
  </div>
<%
Next
If Not hasSubdirs Then
%>
  <div class="fb-empty">No subfolders</div>
<%
End If
Set folder = Nothing
Set fso = Nothing
%>
</div>
<div class="fb-actions">
  <button class="fb-cancel" onclick="closeFolderBrowser()">Cancel</button>
  <form hx-post="/api/playlist.asp" hx-target="#playlist-area" hx-swap="outerHTML"
        onsubmit="closeFolderBrowser()">
    <input type="hidden" name="action" value="save_dir">
    <input type="hidden" name="dir" value="<%= HtmlEsc(browseDir) %>">
    <button type="submit" class="fb-select">Select This Folder</button>
  </form>
</div>
