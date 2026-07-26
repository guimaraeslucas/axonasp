<!--#include file="include/helpers.inc"-->
<%
' ============================================================
' edit.asp - Edit an existing task
' Handles: load task by ID (GET), save changes (POST)
' ============================================================

Dim taskId, taskData, alertMsg, alertType
alertMsg = ""
alertType = ""

taskId = Request.QueryString("id")
If Len(taskId) = 0 Or Not IsNumeric(taskId) Then
    Response.Redirect "default.asp"
End If

' ============================================================
' POST: Save updated task
' ============================================================
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim saveId, saveTitle, saveNotes, savePriority, saveStatus
    saveId       = Request.Form("id")
    saveTitle    = Trim(Request.Form("title"))
    saveNotes    = Trim(Request.Form("notes"))
    savePriority = Trim(Request.Form("priority"))
    saveStatus   = Trim(Request.Form("status"))

    If Len(saveTitle) > 0 Then
        UpdateTask CLng(saveId), saveTitle, saveNotes, savePriority, saveStatus
        Response.Redirect "default.asp?msg=" & Server.URLEncode("Task updated.") & "&type=success"
    Else
        alertMsg  = "Task title cannot be empty."
        alertType = "error"
    End If
End If

' ============================================================
' GET: Load task data
' ============================================================
taskData = GetTask(CLng(taskId))
If IsEmpty(taskData) Then
    Response.Redirect "default.asp?msg=" & Server.URLEncode("Task not found.") & "&type=error"
End If

Dim tId, tTitle, tNotes, tPriority, tStatus, tCreated
tId       = taskData(0)
tTitle    = taskData(1)
tNotes    = taskData(2)
tPriority = taskData(3)
tStatus   = taskData(4)
tCreated  = taskData(5)

%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Edit Task - My To-Do List</title>
<link rel="stylesheet" href="/style.css">
</head>
<body>

<div class="header">
    <div class="wrap">
        <h1>&#9998; Edit Task</h1>
        <div class="subtitle">Update the details below.</div>
    </div>
</div>

<div class="main">

<% If Len(alertMsg) > 0 Then %>
<div class="alert alert-<%= alertType %>"><%= Esc(alertMsg) %></div>
<% End If %>

<div class="edit-form">
    <form method="post" action="edit.asp?id=<%= tId %>">
        <input type="hidden" name="id" value="<%= tId %>">
        <div class="field">
            <label>Title</label>
            <input type="text" name="title" value="<%= Esc(tTitle) %>" required>
        </div>
        <div class="field">
            <label>Notes</label>
            <textarea name="notes" placeholder="Optional notes..."><%= Esc(tNotes) %></textarea>
        </div>
        <div class="field">
            <label>Priority</label>
            <select name="priority">
                <option value="low"    <% If tPriority = "low"    Then %>selected<% End If %>>Low</option>
                <option value="medium" <% If tPriority = "medium" Then %>selected<% End If %>>Medium</option>
                <option value="high"   <% If tPriority = "high"   Then %>selected<% End If %>>High</option>
            </select>
        </div>
        <div class="field">
            <label>Status</label>
            <select name="status">
                <option value="active" <% If tStatus = "active" Then %>selected<% End If %>>Active</option>
                <option value="done"    <% If tStatus = "done"   Then %>selected<% End If %>>Completed</option>
            </select>
        </div>
        <div class="field">
            <label>Created</label>
            <span class="text-muted"><%= Esc(tCreated) %></span>
        </div>
        <div class="actions">
            <button type="submit" class="btn btn-primary">Save Changes</button>
            <a href="default.asp" class="btn btn-secondary">Cancel</a>
        </div>
    </form>
</div>

</div>

<div class="footer">
    <a href="default.asp">&larr; Back to list</a>
</div>

</body>
</html>