<!--#include file="include/helpers.inc"-->
<%
' ============================================================
' default.asp - Main task list page
' Handles: add (POST), toggle complete (GET), delete (GET)
' Displays: filter tabs (all/active/done) + task list
' ============================================================

Dim alertMsg, alertType
alertMsg = ""
alertType = ""

' ============================================================
' POST: Add new task
' ============================================================
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim newTitle, newNotes, newPriority
    newTitle    = Trim(Request.Form("title"))
    newNotes    = Trim(Request.Form("notes"))
    newPriority = Trim(Request.Form("priority"))

    If Len(newPriority) = 0 Then newPriority = "medium"

    If Len(newTitle) > 0 Then
        CreateTask newTitle, newNotes, newPriority
        Response.Redirect "default.asp?msg=" & Server.URLEncode("Task added.") & "&type=success"
    Else
        alertMsg  = "Task title cannot be empty."
        alertType = "error"
    End If
End If

' ============================================================
' GET: Handle actions (toggle / delete)
' ============================================================
Dim action, taskId
action = Request.QueryString("action")
taskId = Request.QueryString("id")

If Len(action) > 0 And Len(taskId) > 0 And IsNumeric(taskId) Then
    Select Case action
        Case "toggle"
            ToggleTask CLng(taskId)
            Response.Redirect "default.asp"

        Case "delete"
            DeleteTask CLng(taskId)
            Response.Redirect "default.asp?msg=" & Server.URLEncode("Task deleted.") & "&type=success"
    End Select
End If

' ============================================================
' GET: Display alert message
' ============================================================
If Len(Request.QueryString("msg")) > 0 Then
    alertMsg  = Request.QueryString("msg")
    alertType = Request.QueryString("type")
End If

' ============================================================
' Load tasks and filter
' ============================================================
Dim filter, allTasks, totalAll, totalActive, totalDone
filter = Request.QueryString("filter")
If Len(filter) = 0 Then filter = "all"

allTasks = GetAllTasks()
totalAll    = UBound(allTasks) + 1
totalActive = CountByStatus(allTasks, "active")
totalDone   = CountByStatus(allTasks, "done")

%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>My To-Do List</title>
<link rel="stylesheet" href="/style.css">
</head>
<body>

<div class="header">
    <div class="wrap">
        <h1>&#9745; My To-Do List</h1>
        <div class="subtitle">Stay organized, one task at a time.</div>
    </div>
</div>

<div class="main">

<% If Len(alertMsg) > 0 Then %>
<div class="alert alert-<%= alertType %>"><%= Esc(alertMsg) %></div>
<% End If %>

<!-- Add Task Form -->
<div class="add-form">
    <form method="post" action="default.asp">
        <div class="row">
            <input type="text" name="title" placeholder="What needs to be done?" autofocus>
            <select name="priority">
                <option value="low">Low</option>
                <option value="medium" selected>Medium</option>
                <option value="high">High</option>
            </select>
            <button type="submit">Add</button>
        </div>
        <div style="margin-top:8px">
            <input type="text" name="notes" placeholder="Optional notes..." style="width:100%;border:1px solid var(--border);border-radius:var(--radius);padding:8px 12px;font-size:13px;font-family:inherit;outline:none">
        </div>
    </form>
</div>

<!-- Filter Tabs -->
<div class="tabs">
    <a href="default.asp?filter=all" class="<% If filter = "all" Then %>active<% End If %>">All<span class="count">(<%= totalAll %>)</span></a>
    <a href="default.asp?filter=active" class="<% If filter = "active" Then %>active<% End If %>">Active<span class="count">(<%= totalActive %>)</span></a>
    <a href="default.asp?filter=done" class="<% If filter = "done" Then %>active<% End If %>">Completed<span class="count">(<%= totalDone %>)</span></a>
</div>

<!-- Task List -->
<% If totalAll = 0 Then %>
<div class="empty">
    <div class="icon">&#9745;</div>
    <p>No tasks yet. Add one above to get started!</p>
</div>

<% Else %>
<ul class="task-list">
    <%
    Dim i, t, showTask
    For i = 0 To UBound(allTasks)
        Set t = allTasks(i)
        showTask = False
        Select Case filter
            Case "all":    showTask = True
            Case "active": showTask = (t(4) = "active")
            Case "done":   showTask = (t(4) = "done")
        End Select

        If showTask Then
    %>
    <li class="task-item <% If t(4) = "done" Then %>done<% End If %>">
        <div class="check">
            <a href="default.asp?action=toggle&id=<%= t(0) %>" title="Toggle complete">&#10003;</a>
        </div>
        <div class="body">
            <div class="title"><%= Esc(t(1)) %></div>
            <% If Len(t(2)) > 0 Then %>
            <div class="notes"><%= Esc(t(2)) %></div>
            <% End If %>
            <div class="meta">
                <span class="badge <%= PriorityClass(t(3)) %>"><%= Esc(t(3)) %></span>
                <span class="text-muted">Created: <%= Esc(t(5)) %></span>
            </div>
        </div>
        <div class="actions">
            <a href="edit.asp?id=<%= t(0) %>">Edit</a>
            <a href="default.asp?action=delete&id=<%= t(0) %>" class="delete"
               onclick="return confirm('Delete this task?');">Delete</a>
        </div>
    </li>
    <%
        End If
    Next
    %>
</ul>
<% End If %>

</div>

<div class="footer">
    Built with AxonHTA &mdash; VBScript + HTML + CSS
</div>

</body>
</html>