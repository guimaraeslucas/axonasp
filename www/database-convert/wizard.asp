<%
' AxonASP Database Export Wizard
' Step-based implementation

Dim Step
Step = Request.QueryString("step")
If Step = "" Then Step = 0

Dim accessPath, sqlitePath, sqliteName
accessPath = Session("AccessPath")
sqliteName = Session("SqliteName")
If sqliteName = "" Then sqliteName = "imported_data.db"
sqlitePath = Server.MapPath(sqliteName)

' Helper to render the header
Sub RenderHeader(title)
%>
<html lang="en">
    <!--
        
        AxonASP Server
        Copyright (C) 2026 G3pix Ltda. All rights reserved.
        
        Developed by Lucas Guimarães - G3pix Ltda
        Contact: https://g3pix.com.br/
        Project URL: https://g3pix.com.br/axonasp
        
        This Source Code Form is subject to the terms of the Mozilla Public
        License, v. 2.0. If a copy of the MPL was not distributed with this
        file, You can obtain one at https://mozilla.org/MPL/2.0/.
        
        Attribution Notice:
        If this software is used in other projects, the name "AxonASP Server"
        must be cited in the documentation or "About" section.
        
        Contribution Policy:
        Modifications to the core source code of AxonASP Server must be
        made available under this same license terms.
        
        -->
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title><%= title %></title>
    <link rel="stylesheet" href="98.css">
    <style>
        body {
            margin: 0;
            font-size: 12px;
            font-family: Tahoma, sans-serif;
            user-select: none;
        }
        *{ font-family: Tahoma, sans-serif!important; }
        .container {
            display: flex;
            flex-direction: column;
            height: 100vh;
            justify-content: space-between;
        }
        .main {
            display: flex;
            flex-grow: 1;
        }
        .sidebar {
            background-color: #060c88;
            min-width: 164px;
            height: 100%;
        }
        .sidebar img {
            width: 100%;
        }
        .content {
            padding: 10px;
            flex-grow: 1;
        }
        .content-inner {
            padding: 2px;
        }
        .content-inner .title {
            margin-bottom: 10px;
            font-size: 1.5rem;
        }
        footer {
            padding: 10px;
            border-top: #9D9D9C ridge 1px;
            text-align: right;
        }
        .buttons button {
            margin-left: 5px;
            min-width: 80px;
        }
        .sunken-panel {
            overflow: auto;
            background: white;
        }
        table.interactive tr.highlighted {
            background-color: #000080;
            color: white;
        }
        ::file-selector-button {
        background: silver;
        border: none;
        border-radius: 0;
        box-shadow: inset -1px -1px #0a0a0a, inset 1px 1px #fff, inset -2px -2px grey, inset 2px 2px #dfdfdf;
        box-sizing: border-box;
        color: transparent;
        min-height: 23px;
        min-width: 75px;
        padding: 0 12px;
        text-shadow: 0 0 #222;
        }
        ::file-selector-button:active {
        box-shadow: inset -1px -1px #fff, inset 1px 1px #0a0a0a, inset -2px -2px #dfdfdf, inset 2px 2px grey;
        text-shadow: 1px 1px #222;
        }
    </style>
</head>
<body>
<div class="container">
    <div class="main">
        <div class="sidebar">
            <img src="server.png" alt="G3Pix AxonASP Database Conversion Tool">
        </div>
        <div class="content">
            <div class="content-inner">
<%
End Sub

' Helper to render the footer
Sub RenderFooter(prevStep, nextStep, nextLabel, nextDisabled)
%>
            </div>
        </div>
    </div>
    <footer>
        <div class="buttons">
            <%
            If prevStep <> "" Then
            %>
                <button onclick="window.location='?step=<%= prevStep %>';">< Back</button>
            <%
            Else
            %>
                <button disabled="disabled">< Back</button>
            <%
            End If
            %>
            
            <%
            If nextStep <> "" Then
            %>
                <button id="btnNext" onclick="<%= nextStep %>" <%= IfThen(nextDisabled, "disabled=""disabled""", "") %> style="margin-left:-3px;"><%= nextLabel %></button>
            <%
            End If
            %>
            
            <!-- <button onclick="if(confirm('Cancel wizard?')) window.close();" style="margin-left:10px;">Cancel</button>
        </div>
    </footer>
</div>
</body>
</html>
<%
End Sub

Function IfThen(cond, t, f)
    If cond Then IfThen = t Else IfThen = f
End Function

' --- Step Logic ---
Select Case Step
Case 0 ' Welcome
    RenderHeader "Welcome"
%>
            <div class="title">AxonASP Database Conversion Tool</div>
            <p>This wizard helps you convert your Microsoft Access database to another modern database format like SQLite, MySQL, PostgreSQL, or SQL Server.</p>
            <p>The system accommodates Access files with a maximum size of 25MB. Once uploaded, you may select the specific tables for transfer. Please note that the system might overwrite existing databases; therefore, we recommend performing a backup prior to the import process."</p>
            <p>Click Next to start.</p>
<%
RenderFooter "", "window.location='?step=1';", "Next >", False

Case 1 ' Select / Upload Database
    ' ... (already updated) ...

    RenderHeader "Select Database"
%>
            <div class="title">Select database</div>
            <p>Choose an Access file and configure the target database:</p>
            
            <form id="uploadForm" action="?step=2" method="post" enctype="multipart/form-data">
                <fieldset>
                    <legend>Source Access Database</legend>
                    <div class="field-row-stacked" style="width: 100%;">
                        <label for="database">Access File (.mdb, .accdb):</label>
                        <input type="file" id="database" name="database" accept=".mdb,.accdb">
                    </div>
                </fieldset>
                <br>
                <fieldset>
                    <legend>Target Database Configuration</legend>
                    <div class="field-row">
                        <label for="dbType">Target Type:</label>
                        <select id="dbType" name="dbType" onchange="updateUI()">
                            <option value="sqlite">SQLite</option>
                            <option value="mysql">MySQL</option>
                            <option value="postgres">PostgreSQL</option>
                            <option value="mssql">MS SQL Server</option>
                        </select>
                    </div>
                    <div class="field-row">
                        <input type="radio" id="modeEnv" name="connMode" value="env" checked onclick="updateUI()">
                        <label for="modeEnv">Use System Default (.env)</label> &nbsp; 
                        <input type="radio" id="modeManual" name="connMode" value="manual" onclick="updateUI()">
                        <label for="modeManual">Manual Configuration</label>
                    </div>
                    
                    <div id="manualFields" style="display:none; margin-top:10px; border-top:1px solid #ccc; padding-top:10px;">
                        <div id="rowHost" class="field-row-stacked">
                            <label for="dbHost">Server/Host:</label>
                            <input type="text" id="dbHost" name="dbHost" value="localhost" placeholder="<%= axgetenv("MYSQL_HOST") %>">
                        </div>
                        <div id="rowPort" class="field-row-stacked">
                            <label for="dbPort">Port:</label>
                            <input type="text" id="dbPort" name="dbPort" value="" placeholder="3306">
                        </div>
                        <div id="rowDB" class="field-row-stacked">
                            <label for="dbName">Database Name / Path:</label>
                            <input type="text" id="dbName" name="dbName" value="imported_data.db" placeholder="<%= axgetenv("MYSQL_DATABASE") %>">
                        </div>
                        <div id="rowUser" class="field-row-stacked">
                            <label for="dbUser">Username:</label>
                            <input type="text" id="dbUser" name="dbUser" value="" placeholder="<%= axgetenv("MYSQL_USER") %>">
                        </div>
                        <div id="rowPass" class="field-row-stacked">
                            <label for="dbPass">Password:</label>
                            <input type="password" id="dbPass" name="dbPass" value="">
                        </div>
                    </div>
                </fieldset>
            </form>
            <script>
                function updateUI() {
                    const type = document.getElementById('dbType').value;
                    const mode = document.querySelector('input[name="connMode"]:checked').value;
                    const manual = document.getElementById('manualFields');
                    
                    if (mode === 'manual') {
                        manual.style.display = 'block';
                        const rowHost = document.getElementById('rowHost');
                        const rowPort = document.getElementById('rowPort');
                        const rowUser = document.getElementById('rowUser');
                        const rowPass = document.getElementById('rowPass');
                        const dbLabel = document.querySelector('label[for="dbName"]');
                        
                        if (type === 'sqlite') {
                            rowHost.style.display = 'none';
                            rowPort.style.display = 'none';
                            rowUser.style.display = 'none';
                            rowPass.style.display = 'none';
                            dbLabel.innerText = "SQLite File Path:";
                        } else {
                            rowHost.style.display = 'block';
                            rowPort.style.display = 'block';
                            rowUser.style.display = 'block';
                            rowPass.style.display = 'block';
                            dbLabel.innerText = "Database Name:";
                            
                            // Set default ports
                            const portInput = document.getElementById('dbPort');
                            if (type === 'mysql') portInput.placeholder = "3306";
                            else if (type === 'postgres') portInput.placeholder = "5432";
                            else if (type === 'mssql') portInput.placeholder = "1433";
                        }
                    } else {
                        manual.style.display = 'none';
                    }
                }
            </script>
<%
RenderFooter "0", "document.getElementById('uploadForm').submit();", "Next >", False

Case 2 ' Process Upload And List Tables
    ' Handle Upload
    Dim uploader, result
    Set uploader = Server.CreateObject("G3FileUploader")
    uploader.MaxFileSize = 50 * 1024 * 1024 ' 50MB
    uploader.AllowExtensions "mdb,accdb"
    uploader.SetUseAllowedOnly True

    Set result = uploader.Process("database", "/temp/uploads")

    ' Save Target Configuration
    Session("DbType") = Request.Form("dbType")
    Session("ConnMode") = Request.Form("connMode")
    Session("DbHost") = Request.Form("dbHost")
    Session("DbPort") = Request.Form("dbPort")
    Session("DbName") = Request.Form("dbName")
    Session("DbUser") = Request.Form("dbUser")
    Session("DbPass") = Request.Form("dbPass")

    If result("IsSuccess") Then
        accessPath = result("FinalPath")
        Session("AccessPath") = accessPath

        ' Validate Target Connection
        Dim dbType, connMode, dbHost, dbPort, dbName, dbUser, dbPass, targetConnStr
        dbType = Session("DbType")
        connMode = Session("ConnMode")
        dbHost = Session("DbHost")
        dbPort = Session("DbPort")
        dbName = Session("DbName")
        dbUser = Session("DbUser")
        dbPass = Session("DbPass")

        If connMode = "env" Then
            Select Case dbType
                Case "sqlite"
                targetConnStr = "Driver={SQLite3};Data Source=" & axgetenv("SQLITE_PATH")
                Case "mysql"
                targetConnStr = "Driver={MySQL};Server=" & axgetenv("MYSQL_HOST") & ";Port=" & axgetenv("MYSQL_PORT") & ";Database=" & axgetenv("MYSQL_DATABASE") & ";uid=" & axgetenv("MYSQL_USER") & ";pwd=" & axgetenv("MYSQL_PASS")
                Case "postgres"
                targetConnStr = "Driver={PostgreSQL};Server=" & axgetenv("POSTGRES_HOST") & ";Port=" & axgetenv("POSTGRES_PORT") & ";Database=" & axgetenv("POSTGRES_DATABASE") & ";uid=" & axgetenv("POSTGRES_USER") & ";pwd=" & axgetenv("POSTGRES_PASS")
                Case "mssql"
                targetConnStr = "Provider=SQLOLEDB;Server=" & axgetenv("MSSQL_HOST") & "," & axgetenv("MSSQL_PORT") & ";Database=" & axgetenv("MSSQL_DATABASE") & ";uid=" & axgetenv("MSSQL_USER") & ";pwd=" & axgetenv("MSSQL_PASS")
            End Select
        Else
            Select Case dbType
                Case "sqlite"
                targetConnStr = "Driver={SQLite3};Data Source=" & Server.MapPath(dbName)
                Case "mysql"
                If dbPort = "" Then dbPort = "3306"
                targetConnStr = "Driver={MySQL};Server=" & dbHost & ";Port=" & dbPort & ";Database=" & dbName & ";uid=" & dbUser & ";pwd=" & dbPass
                Case "postgres"
                If dbPort = "" Then dbPort = "5432"
                targetConnStr = "Driver={PostgreSQL};Server=" & dbHost & ";Port=" & dbPort & ";Database=" & dbName & ";uid=" & dbUser & ";pwd=" & dbPass
                Case "mssql"
                If dbPort = "" Then dbPort = "1433"
                targetConnStr = "Provider=SQLOLEDB;Server=" & dbHost & "," & dbPort & ";Database=" & dbName & ";uid=" & dbUser & ";pwd=" & dbPass
            End Select
        End If

        Dim connTarget, targetOk, targetErr
        targetOk = True
        On Error Resume Next
        Set connTarget = Server.CreateObject("ADODB.Connection")
        connTarget.Open targetConnStr
        If Err.Number <> 0 Then
            targetOk = False
            targetErr = Err.Description
        Else
            connTarget.Close
        End If
        On Error Goto 0

        If Not targetOk Then
            RenderHeader "Target Connection Error"
%>
                    <div class="title"><img src="error.png" style="vertical-align: middle;"> Target Connection Error</div>
                    <p>Could not connect to the destination database (<b><%= dbType %></b>):</p>
                    <div id="log" class="field-border" style="padding: 8px; height: 150px; width: 95%; margin-top: 10px; overflow: auto;overflow-wrap: anywhere;">
                    <%= targetErr %>
                    </div>
                    <p>Please go back and verify your settings.</p>
<%
RenderFooter "1", "", "Next >", True
Else
    ' Connection OK, show tables
    RenderHeader "Database tables"
%>
                <style>
                input[type=checkbox], input[type=radio] {
                appearance: auto !important;
                background: #FFF !important;
                border: solid 1px #000 !important;
                margin: 0px 5px 0px 0px !important;
                opacity: 1 !important;
                position: relative !important;
                vertical-align: text-bottom;}
                </style>
                <div class="title">Database tables</div>
                <p>These are the tables found in the Access database. Select the ones you want to import:</p>
                <div class="sunken-panel" style="height: 200px; width: 100%;">
                    <table class="interactive" id="tblTables" style="width: 100%;">
                        <thead>
                            <tr>
                                <th style="text-align: left; padding-left: 5px;">
                                    <input type="checkbox" id="chkAll" name="chkAll" checked onclick="toggleAll(this)"> Table Name
                                </th>
                            </tr>
                        </thead>
                        <tbody>
<%
' List Tables using ADOX
Dim cat, tbl, Count
Count = 0
On Error Resume Next
Set cat = Server.CreateObject("ADOX.Catalog")
cat.ActiveConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessPath

If Err.Number <> 0 Then
    Err.Clear
    cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & accessPath
End If

If Err.Number <> 0 Then
%>
                    <tr><td style="color:red;">Error processing database: <%= Err.Description %></td></tr>
<%
Else
    For Each tbl In cat.Tables
        If tbl.Type = "TABLE" Or tbl.Type = "LINK" Then
            Count = Count + 1
%>
                            <tr>
                                <td>
                                    <input type="checkbox" name="tables" value="<%= tbl.Name %>" checked>
                                    <%= tbl.Name %>
                                </td>
                            </tr>
<%
End If
Next
End If
On Error Goto 0
%>
                        </tbody>
                    </table>
                </div>
                <form id="importForm" action="?step=3" method="post">
                    <input type="hidden" id="selectedTables" name="selectedTables">
                </form>
                <script>
                    function toggleAll(chk) {
                        const checks = document.querySelectorAll('input[name="tables"]');
                        checks.forEach(c => c.checked = chk.checked);
                    }
                    function prepareImport() {
                        const selected = [];
                        document.querySelectorAll('input[name="tables"]:checked').forEach(c => {
                            selected.push(c.value);
                        });
                        if (selected.length === 0) {
                            alert('Please select at least one table.');
                            return;
                        }
                        document.getElementById('selectedTables').value = selected.join('|');
                        document.getElementById('importForm').submit();
                    }
                </script>
<%
RenderFooter "1", "prepareImport();", "Next >", (Count = 0)
End If
Else
    RenderHeader "Upload Error"
%>
                <div class="title"><img src="error.png" style="vertical-align: middle;"> Upload Error</div>
                <div id="log" class="field-border" style="padding: 8px; height: 150px; width: 95%; margin-top: 10px; overflow: auto;overflow-wrap: anywhere;">
                <%= result("ErrorMessage") %>
                </div>
<%
RenderFooter "1", "", "Next >", True
End If

Case 3 ' Import Execution
    Dim selectedTables
    selectedTables = Request.Form("selectedTables")
    If selectedTables = "" Then Response.Redirect "?step=2"

    Session("SelectedTables") = selectedTables

    RenderHeader "Importing"
%>
            <div class="title">Importing</div>
            <p id="statusMsg" style="overflow-wrap: anywhere;"><img src="hourglass.gif" style="vertical-align: middle;"> Preparing import... </p>
            <div class="progress-indicator segmented" style="height: 24px;">
                <span class="progress-indicator-bar" id="progressBar" style="width: 0%;"></span>
            </div>
           <!-- <div id="log" class="sunken-panel" style="height: 150px; width: 100%; margin-top: 10px; font-family: tahoma, verdana, arial, sans-serif; font-size: 12pt; padding: 5px;">-->
           <div id="log" class="field-border" style="padding: 8px; height: 150px; width: 95%; margin-top: 10px; overflow: auto;overflow-wrap: anywhere;">
            </div>

            <script>
                const tables = "<%= selectedTables %>".split('|');
                let currentIndex = 0;
                
                function updateProgress(percent, msg) {
                    document.getElementById('progressBar').style.width = percent + '%';
                    if (msg) {
                        document.getElementById('statusMsg').innerText = msg;
                        const log = document.getElementById('log');
                        log.innerHTML += '<div>' + msg + '</div>';
                        log.scrollTop = log.scrollHeight;
                    }
                }

                function doNextTable() {
                    if (currentIndex >= tables.length) {
                        updateProgress(100, "Import completed successfully!");
                        document.getElementById('btnNext').disabled = false;
                        return;
                    }

                    const tableName = tables[currentIndex];
                    updateProgress((currentIndex / tables.length) * 100, "Importing table: " + tableName + "...");
                    
                    fetch('import_worker.asp?table=' + encodeURIComponent(tableName))
                        .then(response => response.text())
                        .then(text => {
                            if (text.indexOf('SUCCESS') !== -1) {
                                currentIndex++;
                                doNextTable();
                            } else {
                                updateProgress((currentIndex / tables.length) * 100, "Error importing " + tableName + ": " + text);
                                document.getElementById('btnNext').disabled = false;
                                document.getElementById('btnNext').innerText = "Finish";
                            }
                        })
                        .catch(err => {
                            updateProgress(0, "Fetch error: " + err);
                        });
                }

                window.onload = function() {
                    // Disable next button until finished
                    document.getElementById('btnNext').disabled = true;
                    doNextTable();
                };
            </script>
<%
RenderFooter "2", "window.location='?step=4';", "Next >", True

Case 4 ' Success
    RenderHeader "Success"
    Dim targetInfo
    If Session("ConnMode") = "env" Then
        targetInfo = "System Default (" & Session("DbType") & ")"
    Else
        targetInfo = Session("DbName") & " (" & Session("DbType") & ")"
    End If
%>
            <div class="title">Success</div>
            <p>Database imported with success to <strong><%= targetInfo %></strong></p>
            <p>You may close this window now.</p>
<%
RenderFooter "0", "window.top.location='/';", "Close", False
End Select
%>
