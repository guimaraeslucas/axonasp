<%
Option Explicit
Response.ContentType = "text/html; charset=utf-8"
%>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>ASP Project Builder - AxonASP Code Generation Assistant</title>
        <link rel="stylesheet" href="../css/axonasp.css" />
        <style>
            html,
            body {
                width: 100%;
                height: 100%;
                margin: 0;
                padding: 0;
            }

            body {
                font-family: Tahoma, Verdana, Arial, sans-serif;
                background-color: #ece9d8;
                color: #000;
                overflow: hidden;
            }

            #main-container {
                display: flex;
                width: 100%;
                max-width: none !important;
                margin: 0;
                padding: 0 !important;
                height: calc(100vh - 63px);
            }

            .sidebar {
                width: 280px;
                background-color: #e2e2e2;
                border-right: 2px solid #aca899;
                overflow-y: auto;
                padding: 15px;
                font-size: 12px;
            }

            .sidebar h3 {
                font-size: 12px;
                font-weight: bold;
                color: #003399;
                border-bottom: 1px solid #aca899;
                padding-bottom: 5px;
                margin: 10px 0 8px 0;
            }

            .sidebar ul {
                list-style: none;
                margin: 0;
                padding: 0 0 0 10px;
            }

            .sidebar li {
                margin: 4px 0;
            }

            .sidebar a {
                color: #000;
                text-decoration: none;
                display: block;
                padding: 2px 4px;
            }

            .sidebar a:hover {
                color: #0000ff;
                text-decoration: underline;
            }

            #content {
                flex: 1;
                background-color: #fff;
                overflow-y: auto;
                padding: 25px 40px;
            }

            #content h1 {
                color: #000;
                font-size: 24px;
                border-bottom: 2px solid #3366cc;
                padding-bottom: 5px;
                margin-top: 0;
                margin-bottom: 8px;
            }

            #content h2 {
                color: #003399;
                font-size: 14px;
                font-weight: bold;
                margin-top: 20px;
                margin-bottom: 8px;
                border-bottom: 1px solid #e2e2e2;
                padding-bottom: 3px;
            }

            .intro-text {
                font-size: 12px;
                color: #333;
                margin-bottom: 15px;
                line-height: 1.5;
            }

            .form-section {
                background-color: #f5f5f5;
                border: 1px solid #aca899;
                padding: 12px;
                margin-bottom: 15px;
                border-radius: 0;
            }

            .form-section p {
                font-size: 11px;
                color: #555;
                margin: 0 0 8px 0;
                line-height: 1.4;
            }

            .form-input-area {
                margin-bottom: 10px;
            }

            .form-input-area label {
                display: block;
                font-size: 11px;
                font-weight: bold;
                margin-bottom: 4px;
                color: #003399;
            }

            textarea,
            input[type="text"],
            select {
                width: 100%;
                box-sizing: border-box;
                border: 1px solid #aca899;
                padding: 6px;
                font-family: Tahoma, Verdana, Arial, sans-serif;
                font-size: 11px;
                background-color: #fff;
            }

            textarea {
                resize: vertical;
                min-height: 100px;
            }

            .radio-group,
            .checkbox-group {
                margin: 8px 0;
            }

            .radio-item,
            .check-item {
                display: inline-block;
                margin-right: 20px;
                font-size: 11px;
            }

            .radio-item input[type="radio"],
            .check-item input[type="checkbox"] {
                margin-right: 4px;
                cursor: pointer;
            }

            .radio-item label,
            .check-item label {
                cursor: pointer;
            }

            .radio-item.selected label,
            .check-item.selected label {
                font-weight: bold;
                color: #003399;
            }

            .two-col {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 15px;
                margin-bottom: 15px;
            }

            .tag {
                font-size: 9px;
                color: #666;
                font-weight: normal;
            }

            .button-group {
                display: flex;
                gap: 8px;
                margin: 20px 0;
            }

            .btn {
                padding: 8px 16px;
                border: 2px solid #aca899;
                background-color: #e2e2e2;
                color: #000;
                cursor: pointer;
                font-family: Tahoma, Verdana, Arial, sans-serif;
                font-size: 11px;
                font-weight: bold;
            }

            .btn:hover {
                background-color: #d0d0d0;
                border-color: #808080;
            }

            .btn-primary {
                background-color: #003399;
                color: white;
                border-color: #003399;
            }

            .btn-primary:hover {
                background-color: #2952a3;
                border-color: #2952a3;
            }

            #output-area {
                background-color: #f5f5f5;
                border: 2px solid #aca899;
                padding: 12px;
                margin-top: 20px;
                display: none;
            }

            #output-area.show {
                display: block;
            }

            #output-area h3 {
                margin-top: 0;
                color: #003399;
                font-size: 12px;
                border-bottom: 1px solid #aca899;
                padding-bottom: 5px;
            }

            #output-textarea {
                width: 100%;
                height: 300px;
                border: 1px solid #aca899;
                padding: 8px;
                font-family: "Courier New", Monospace;
                font-size: 10px;
                background-color: #fff;
                box-sizing: border-box;
            }

            .copy-toast {
                position: fixed;
                bottom: 20px;
                right: 20px;
                background-color: #003399;
                color: white;
                padding: 12px 20px;
                border: 1px solid #000;
                font-size: 11px;
                opacity: 0;
                transition: opacity 0.3s;
                pointer-events: none;
                z-index: 1000;
            }

            .copy-toast.show {
                opacity: 1;
                pointer-events: auto;
            }

            .note {
                font-size: 10px;
                color: #666;
                font-style: italic;
                margin-top: 4px;
            }
        </style>
    </head>
    <body>
        <div id="header">
            <div class="logo">
                <img
                    src="data:image/png;base64,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"
                    alt="AxonASP"
                    width="43"
                />
            </div>
            <h1>ASP Code Generator</h1>
        </div>

        <div id="main-container">
            <div class="sidebar" id="sidebar">
                <div class="section-title">User Guide</div>
                <ul>
                    <li><a href="#section-app">Application Details</a></li>
                    <li><a href="#section-arch">Architecture</a></li>
                    <li><a href="#section-features">Features</a></li>
                    <li><a href="#section-lang">Localization</a></li>
                    <li><a href="#section-extra">Advanced Options</a></li>
                    <li><a href="#section-output">Generated Output</a></li>
                </ul>

                <div class="section-title">About ASP Builder</div>
                <p style="font-size: 11px; color: #555; margin: 8px 0">
                    This tool generates structured prompts for AI coding agents
                    building applications with AxonASP and ASP Classic. Provide
                    your requirements and receive a comprehensive markdown
                    document ready for your chosen agent.
                </p>

                <div class="section-title">AxonASP Resources</div>
                <ul>
                    <li><a href="/">Home</a></li>
                    <li><a href="/manual/">Documentation</a></li>
                </ul>
            </div>

            <div id="content">
                <h1>Code Generation Assistant</h1>
                <p class="intro-text">
                    Build structured AI prompts for developing web applications
                    using AxonASP. Describe your vision, choose your
                    preferences, and generate a comprehensive guideline
                    document. Perfect for collaboration with development agents
                    and teams.
                </p>

                <!-- Application Details Section -->
                <div id="section-app">
                    <h2>Core Application Definition</h2>
                    <div class="form-section">
                        <h3
                            style="
                                margin-top: 0;
                                color: #003399;
                                font-size: 12px;
                                border: none;
                            "
                        >
                            Project Name
                        </h3>
                        <p>Assign a name for your application.</p>
                        <div class="form-input-area">
                            <input
                                type="text"
                                id="appname"
                                placeholder="Example: My Blog"
                            />
                        </div>
                    </div>
                    <div class="form-section">
                        <h3
                            style="
                                margin-top: 0;
                                color: #003399;
                                font-size: 12px;
                                border: none;
                            "
                        >
                            Requirements Statement
                        </h3>
                        <p>
                            Describe the application concept, target users, and
                            primary functionality.
                        </p>
                        <div class="form-input-area">
                            <textarea
                                id="description"
                                placeholder="Example: A time-tracking platform for freelancers to log billable hours, organize projects by client, generate invoices, track expenses per project, and create monthly billing reports."
                            ></textarea>
                        </div>
                        <p class="note">
                            Be specific about core features, workflows, and
                            business logic.
                        </p>
                    </div>
                </div>

                <!-- Architecture Section -->
                <div id="section-arch">
                    <h2>Architecture & Technology Stack</h2>

                    <div class="form-section">
                        <h3
                            style="
                                margin-top: 0;
                                color: #003399;
                                font-size: 12px;
                                border: none;
                            "
                        >
                            Design Pattern
                        </h3>
                        <p>
                            Selection determines code organization and
                            development workflow.
                        </p>
                        <div class="radio-group">
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="style"
                                    value="mvc"
                                    id="style-mvc"
                                    checked
                                />
                                <label for="style-mvc"
                                    >Model-View-Controller</label
                                >
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="style"
                                    value="mvvm"
                                    id="style-mvvm"
                                />
                                <label for="style-mvvm"
                                    >Model-View-ViewModel</label
                                >
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="style"
                                    value="mixed"
                                    id="style-mixed"
                                />
                                <label for="style-mixed">Inline Code</label>
                            </div>
                        </div>
                        <p class="note">
                            MVC recommended for larger projects. Inline suitable
                            for simple pages.
                        </p>
                    </div>

                    <div class="form-section">
                        <h3
                            style="
                                margin-top: 0;
                                color: #003399;
                                font-size: 12px;
                                border: none;
                            "
                        >
                            Data Persistence
                        </h3>
                        <p>
                            Backend database technology for application data
                            storage.
                        </p>
                        <div class="radio-group">
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="database"
                                    value="sqlite"
                                    id="db-sqlite"
                                    checked
                                />
                                <label for="db-sqlite">SQLite</label>
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="database"
                                    value="mysql"
                                    id="db-mysql"
                                />
                                <label for="db-mysql">MySQL</label>
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="database"
                                    value="postgresql"
                                    id="db-psql"
                                />
                                <label for="db-psql">PostgreSQL</label>
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="database"
                                    value="mssql"
                                    id="db-mssql"
                                />
                                <label for="db-mssql">MS SQL Server</label>
                            </div>
                        </div>
                    </div>

                    <div class="form-section">
                        <h3
                            style="
                                margin-top: 0;
                                color: #003399;
                                font-size: 12px;
                                border: none;
                            "
                        >
                            User Interface Framework
                        </h3>
                        <p>
                            CSS framework for responsive design and component
                            styling.
                        </p>
                        <div class="radio-group">
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="ui"
                                    value="axonasp"
                                    id="ui-axonasp"
                                />
                                <label for="ui-axonasp">AxonASP Native</label>
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="ui"
                                    value="bootstrap"
                                    id="ui-bootstrap"
                                />
                                <label for="ui-bootstrap">Bootstrap 5</label>
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="ui"
                                    value="tailwind"
                                    id="ui-tailwind"
                                    checked
                                />
                                <label for="ui-tailwind">Tailwind CSS</label>
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="ui"
                                    value="None"
                                    id="ui-none"
                                />
                                <label for="ui-none">None</label>
                            </div>
                        </div>
                        <div
                            id="axonasp-native-hint"
                            style="
                                display: none;
                                margin-top: 10px;
                                padding: 8px 10px;
                                background-color: #eef0ff;
                                border: 1px solid #335ea8;
                                border-left: 4px solid #003399;
                                font-size: 10px;
                                color: #333;
                                line-height: 1.6;
                            "
                        >
                            <strong style="color: #003399"
                                >AxonASP Native Style Directives:</strong
                            ><br />
                            Retro Microsoft MSDN Era (2003-2005) / Windows XP.
                            Tahoma/Verdana (Primary), arial, helvetica,
                            sans-serif (Fallback). Bold titles with a solid blue
                            <code>border-bottom</code>. Never use emojis or
                            icons that do not fit the era. Header: Linear
                            gradient <code>#003399</code> to
                            <code>#3366CC</code>. Background:
                            <code>#ECE9D8</code> (Beige-gray). Highlight:
                            <code>#335EA8</code>. Borders: Metallic gray
                            <code>#808080</code>. Visual hard-edges.
                            <code>border-radius: 0 !important</code>. Perfectly
                            square inputs/buttons.
                        </div>
                    </div>

                    <div class="form-section">
                        <h3
                            style="
                                margin-top: 0;
                                color: #003399;
                                font-size: 12px;
                                border: none;
                            "
                        >
                            JavaScript Framework
                        </h3>
                        <p>
                            Client-side JavaScript framework for interactivity
                            and dynamic behavior.
                        </p>
                        <div class="radio-group">
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="jsframework"
                                    value="Vanilla JavaScript"
                                    id="js-vanilla"
                                    checked
                                />
                                <label for="js-vanilla"
                                    >Vanilla JavaScript</label
                                >
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="jsframework"
                                    value="jQuery"
                                    id="js-jquery"
                                />
                                <label for="js-jquery">jQuery</label>
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="jsframework"
                                    value="React"
                                    id="js-react"
                                />
                                <label for="js-react">React</label>
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="jsframework"
                                    value="Vue"
                                    id="js-vue"
                                />
                                <label for="js-vue">Vue</label>
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="jsframework"
                                    value="Next.js"
                                    id="js-nextjs"
                                />
                                <label for="js-nextjs">Next.js</label>
                            </div>
                            <div class="radio-item">
                                <input
                                    type="radio"
                                    name="jsframework"
                                    value="Angular"
                                    id="js-angular"
                                />
                                <label for="js-angular">Angular</label>
                            </div>
                        </div>
                    </div>

                    <div class="form-section">
                        <h3
                            style="
                                margin-top: 0;
                                color: #003399;
                                font-size: 12px;
                                border: none;
                            "
                        >
                            Visual Appearance
                        </h3>
                        <div class="two-col">
                            <div>
                                <p
                                    style="
                                        margin: 0 0 8px 0;
                                        font-weight: bold;
                                        color: #003399;
                                        font-size: 11px;
                                    "
                                >
                                    Color Palette
                                </p>
                                <div
                                    class="radio-group"
                                    style="
                                        display: flex;
                                        flex-direction: column;
                                        gap: 6px;
                                    "
                                >
                                    <div class="radio-item" style="margin: 0">
                                        <input
                                            type="radio"
                                            name="color"
                                            value="light"
                                            id="color-light"
                                            checked
                                        />
                                        <label for="color-light"
                                            >Light Theme</label
                                        >
                                    </div>
                                    <div class="radio-item" style="margin: 0">
                                        <input
                                            type="radio"
                                            name="color"
                                            value="dark"
                                            id="color-dark"
                                        />
                                        <label for="color-dark"
                                            >Dark Theme</label
                                        >
                                    </div>
                                    <div class="radio-item" style="margin: 0">
                                        <input
                                            type="radio"
                                            name="color"
                                            value="auto"
                                            id="color-auto"
                                        />
                                        <label for="color-auto"
                                            >Auto-Detect</label
                                        >
                                    </div>
                                </div>
                            </div>
                            <div>
                                <p
                                    style="
                                        margin: 0 0 8px 0;
                                        font-weight: bold;
                                        color: #003399;
                                        font-size: 11px;
                                    "
                                >
                                    Interactive Elements
                                </p>
                                <div
                                    class="radio-group"
                                    style="
                                        display: flex;
                                        flex-direction: column;
                                        gap: 6px;
                                    "
                                >
                                    <div class="radio-item" style="margin: 0">
                                        <input
                                            type="radio"
                                            name="emoji"
                                            value="enabled"
                                            id="emoji-enabled"
                                            checked
                                        />
                                        <label for="emoji-enabled"
                                            >Include Icons</label
                                        >
                                    </div>
                                    <div class="radio-item" style="margin: 0">
                                        <input
                                            type="radio"
                                            name="emoji"
                                            value="minimal"
                                            id="emoji-minimal"
                                        />
                                        <label for="emoji-minimal"
                                            >Minimal Icons</label
                                        >
                                    </div>
                                    <div class="radio-item" style="margin: 0">
                                        <input
                                            type="radio"
                                            name="emoji"
                                            value="none"
                                            id="emoji-none"
                                        />
                                        <label for="emoji-none"
                                            >Text Only</label
                                        >
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Features Section -->
                <div id="section-features">
                    <h2>Functional Components</h2>

                    <div class="form-section">
                        <h3
                            style="
                                margin-top: 0;
                                color: #003399;
                                font-size: 12px;
                                border: none;
                            "
                        >
                            Select Enabled Capabilities
                        </h3>
                        <p>
                            Choose which features to include in your generated
                            specification.
                        </p>
                        <div class="checkbox-group">
                            <div class="check-item">
                                <input type="checkbox" id="feat-auth" checked />
                                <label for="feat-auth"
                                    >User Authentication</label
                                >
                            </div>
                            <div class="check-item">
                                <input type="checkbox" id="feat-sample" />
                                <label for="feat-sample"
                                    >Sample Data Initialization</label
                                >
                            </div>
                            <div class="check-item">
                                <input type="checkbox" id="feat-crud" checked />
                                <label for="feat-crud"
                                    >Data Management (CRUD)</label
                                >
                            </div>
                            <br />
                            <div class="check-item">
                                <input type="checkbox" id="feat-search" />
                                <label for="feat-search"
                                    >Search Capability</label
                                >
                            </div>
                            <div class="check-item">
                                <input type="checkbox" id="feat-upload" />
                                <label for="feat-upload">File Handling</label>
                            </div>
                            <div class="check-item">
                                <input type="checkbox" id="feat-dash" />
                                <label for="feat-dash">Dashboard Page</label>
                            </div>
                            <br />
                            <div class="check-item">
                                <input type="checkbox" id="feat-pdf" />
                                <label for="feat-pdf">PDF Export</label>
                            </div>
                            <div class="check-item">
                                <input type="checkbox" id="feat-email" />
                                <label for="feat-email"
                                    >Email Integration</label
                                >
                            </div>
                            <div class="check-item">
                                <input type="checkbox" id="feat-api" />
                                <label for="feat-api">REST API Endpoints</label>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Localization Section -->
                <div id="section-lang">
                    <h2>Localization & Content Language</h2>

                    <div class="form-section">
                        <h3
                            style="
                                margin-top: 0;
                                color: #003399;
                                font-size: 12px;
                                border: none;
                            "
                        >
                            Application Language
                        </h3>
                        <p>
                            Primary language for all user-facing content and
                            interface text.
                        </p>
                        <div class="form-input-area">
                            <select id="lang-select">
                                <option value="English">English</option>
                                <option value="Brazilian Portuguese">
                                    Brazilian Portuguese
                                </option>
                                <option value="Portuguese">Portuguese</option>
                                <option value="Spanish">Spanish</option>
                                <option value="French">French</option>
                                <option value="German">German</option>
                                <option value="Italian">Italian</option>
                                <option value="Russian">Russian</option>
                                <option value="Japanese">Japanese</option>
                                <option value="Chinese">
                                    Chinese (Simplified)
                                </option>
                                <option value="Chinese (Traditional)">
                                    Chinese (Traditional)
                                </option>
                                <option value="Korean">Korean</option>
                                <option value="Arabic">Arabic</option>
                                <option value="Hindi">Hindi</option>
                                <option value="Vietnamese">Vietnamese</option>
                                <option value="Thai">Thai</option>
                                <option value="Indonesian">Indonesian</option>
                            </select>
                        </div>
                    </div>

                    <div class="form-section">
                        <h3
                            style="
                                margin-top: 0;
                                color: #003399;
                                font-size: 12px;
                                border: none;
                            "
                        >
                            Multi-Language Support
                        </h3>
                        <p>
                            Enable translation framework for content in
                            additional languages.
                        </p>
                        <div class="check-item">
                            <input type="checkbox" id="multilingual" />
                            <label for="multilingual"
                                >Enable Multi-Language Interface</label
                            >
                        </div>
                        <div
                            id="multi-langs-list"
                            style="
                                display: none;
                                margin-top: 10px;
                                padding-top: 10px;
                                border-top: 1px solid #aca899;
                            "
                        >
                            <p
                                style="
                                    font-size: 11px;
                                    margin: 0 0 8px 0;
                                    color: #555;
                                "
                            >
                                Additional language support (in addition to
                                selected primary):
                            </p>
                            <div
                                class="checkbox-group"
                                id="multi-lang-checks"
                            ></div>
                        </div>
                        <p class="note">
                            Framework will include language switching and
                            translation management.
                        </p>
                    </div>
                </div>

                <!-- Advanced Options Section -->
                <div id="section-extra">
                    <h2>Advanced Configuration</h2>

                    <div class="form-section">
                        <h3
                            style="
                                margin-top: 0;
                                color: #003399;
                                font-size: 12px;
                                border: none;
                            "
                        >
                            Custom Requirements
                            <span class="tag">(optional)</span>
                        </h3>
                        <p>
                            Specify additional constraints, architectural
                            decisions, or special implementation details.
                        </p>
                        <div class="form-input-area">
                            <textarea
                                id="extra"
                                placeholder="Example: Implement role-based access control with Admin/User roles. Use stored procedures for complex queries. Include activity logging for all data modifications."
                            ></textarea>
                        </div>
                    </div>
                </div>

                <!-- Output Section -->
                <div id="section-output">
                    <h2>Generated Documentation</h2>
                    <div class="button-group">
                        <button
                            class="btn btn-primary"
                            onclick="generatePrompt()"
                        >
                            Prepare Markdown Document
                        </button>
                        <button class="btn" onclick="copyToClipboard()">
                            Copy Instructions
                        </button>
                    </div>

                    <div id="output-area">
                        <h3>Your Generated Agent Prompt</h3>
                        <p
                            style="
                                font-size: 10px;
                                color: #666;
                                margin: 2px 0 8px 0;
                            "
                        >
                            Copy this markdown and paste it into your AI coding
                            agent.
                        </p>
                        <textarea id="output-textarea" readonly></textarea>
                    </div>
                </div>
            </div>
        </div>

        <div class="copy-toast" id="copy-toast">
            Copied to clipboard successfully!
        </div>
        <script>
            // Helper functions
            function val(id) {
                return (document.getElementById(id) || {}).value || "";
            }
            function chk(id) {
                return (document.getElementById(id) || {}).checked || false;
            }
            function getRadio(name) {
                var el = document.querySelector(
                    'input[name="' + name + '"]:checked'
                );
                return el ? el.value : "";
            }

            // Initialize multi-language support
            function initMultiLang() {
                var langs = [
                    "Brazilian Portuguese",
                    "Portuguese",
                    "Spanish",
                    "French",
                    "German",
                    "Italian",
                    "Russian",
                    "Japanese",
                    "Chinese (Traditional)",
                    "Korean",
                    "Arabic",
                    "Hindi",
                    "Vietnamese",
                    "Thai",
                    "Indonesian",
                ];

                var container = document.getElementById("multi-lang-checks");
                if (container) {
                    container.innerHTML = "";
                    langs.forEach(function (lang) {
                        var id =
                            "ml-" +
                            lang.replace(/[^a-zA-Z]/g, "").toLowerCase();
                        var div = document.createElement("div");
                        div.className = "check-item";
                        div.innerHTML =
                            '<input type="checkbox" id="' +
                            id +
                            '" data-lang="' +
                            lang +
                            '">' +
                            '<label for="' +
                            id +
                            '">' +
                            lang +
                            "</label>";
                        container.appendChild(div);
                    });
                }
            }

            // Multi-language toggle
            if (document.getElementById("multilingual")) {
                document
                    .getElementById("multilingual")
                    .addEventListener("change", function () {
                        var list = document.getElementById("multi-langs-list");
                        if (list) {
                            list.style.display = this.checked
                                ? "block"
                                : "none";
                        }
                    });
            }

            // Show/hide AxonASP Native hint
            document
                .querySelectorAll('input[name="ui"]')
                .forEach(function (radio) {
                    radio.addEventListener("change", function () {
                        var hint = document.getElementById(
                            "axonasp-native-hint"
                        );
                        if (hint) {
                            hint.style.display =
                                this.value === "axonasp" ? "block" : "none";
                        }
                    });
                });

            // Radio styling
            document
                .querySelectorAll(".radio-item input, .check-item input")
                .forEach(function (input) {
                    input.addEventListener("change", function () {
                        var parent =
                            this.closest(".radio-item") ||
                            this.closest(".check-item");
                        if (parent) {
                            if (this.type === "radio") {
                                var group = this.getAttribute("name");
                                document
                                    .querySelectorAll(
                                        'input[name="' + group + '"]'
                                    )
                                    .forEach(function (el) {
                                        var p = el.closest(".radio-item");
                                        if (p) p.classList.remove("selected");
                                    });
                            }
                            parent.classList.toggle("selected", this.checked);
                        }
                    });

                    // Initial styling
                    var parent =
                        input.closest(".radio-item") ||
                        input.closest(".check-item");
                    if (parent && input.checked) {
                        parent.classList.add("selected");
                    }
                });

            function generatePrompt() {
                var desc = val("description").trim();
                if (!desc) {
                    alert("Please provide an application description.");
                    document.getElementById("description").focus();
                    return;
                }

                var features = [];
                if (chk("feat-auth"))
                    features.push(
                        "Secure user authentication with session management"
                    );
                if (chk("feat-sample"))
                    features.push("Pre-populated sample data for testing");
                if (chk("feat-crud"))
                    features.push(
                        "Complete data management (create, read, update, delete operations)-CRUD"
                    );
                if (chk("feat-search"))
                    features.push(
                        "Data search and advanced filtering capabilities"
                    );
                if (chk("feat-upload"))
                    features.push("File upload and attachment handling");
                if (chk("feat-dash"))
                    features.push(
                        "Administrative dashboard with key statistics"
                    );
                if (chk("feat-pdf"))
                    features.push("PDF document generation and export");
                if (chk("feat-email"))
                    features.push("Email integration and notification system");
                if (chk("feat-api"))
                    features.push(
                        "RESTful API endpoints for external integration"
                    );

                var primaryLang = val("lang-select");
                var isMulti = chk("multilingual");
                var supportedLangs = [primaryLang];

                if (isMulti) {
                    document
                        .querySelectorAll("#multi-lang-checks input:checked")
                        .forEach(function (el) {
                            var lang = el.getAttribute("data-lang");
                            if (lang && supportedLangs.indexOf(lang) === -1) {
                                supportedLangs.push(lang);
                            }
                        });
                }

                var style = getRadio("style");
                var db = getRadio("database");
                var ui = getRadio("ui");
                var jsframework = getRadio("jsframework");
                var color = getRadio("color");
                var emoji = getRadio("emoji");

                var md =
                    "# " +
                    (val("appname") || "Web Application") +
                    " - AxonASP Development Specification\n\n";

                md += "## Project Overview\n\n";
                md += desc + "\n\n";

                md += "## Technical Requirements\n\n";

                md += "### Platform & Language\n";
                md += "- **Language:** Classic ASP (VBScript)\n";
                md += "- **Runtime:** AxonASP Virtual Machine\n\n";

                md += "### Architecture\n";
                md += "- **Pattern:** " + style.toUpperCase() + "\n";
                md +=
                    "- **Database:** " +
                    db.charAt(0).toUpperCase() +
                    db.slice(1) +
                    "\n";
                md +=
                    "- **UI Framework:** " +
                    ui.charAt(0).toUpperCase() +
                    ui.slice(1) +
                    "\n";
                if (ui === "axonasp") {
                    md +=
                        "  - _Style Directives:_ Retro Microsoft MSDN Era (2003-2005) / Windows XP. Tahoma/Verdana (Primary), arial, helvetica, sans-serif (Fallback). Bold titles with a solid blue border-bottom. Never use emojis or icons that do not fit the era. Header: Linear gradient #003399 to #3366CC. Background: #ECE9D8 (Beige-gray). Highlight: #335EA8. Borders: Metallic gray #808080. Visual hard-edges. border-radius: 0 !important. Perfectly square inputs/buttons.\n";
                }
                md += "- **JavaScript Framework:** " + jsframework + "\n";
                md +=
                    "- **Theme:** " +
                    color.charAt(0).toUpperCase() +
                    color.slice(1) +
                    "\n";
                md +=
                    "- **Icons:** " +
                    emoji.charAt(0).toUpperCase() +
                    emoji.slice(1) +
                    "\n\n";

                md += "### Supported Languages\n";
                md += "- Primary: " + primaryLang + "\n";
                if (isMulti && supportedLangs.length > 1) {
                    md +=
                        "- Additional: " +
                        supportedLangs.slice(1).join(", ") +
                        "\n";
                }
                md +=
                    "- All UI text must be translatable and centralized. English is always supported.\n\n";

                if (features.length > 0) {
                    md += "### Required Features\n\n";
                    features.forEach(function (feat) {
                        md += "- " + feat + "\n";
                    });
                    md += "\n";
                }

                if (val("extra").trim()) {
                    md += "### Custom Requirements\n\n";
                    md += val("extra").trim() + "\n\n";
                }

                md += "## VBScript & Classic ASP Coding Standards\n\n";

                md += "### Code Structure & Directives\n";
                md +=
                    "1. **Page Directives:** Always include `<%
@Language = "VBSCRIPT"
%>` on the first line\n";
                md +=
                    "2. **Option Explicit:** Enforce strict variable declaration with `Option Explicit`\n";
                md +=
                    "3. **Code Blocks:** **NEVER** use single-line If statements; always use block syntax with End If\n eg: ```\nIf condition Then\n    ' code\nEnd If\n```\n This is critical and common source of bugs. Be sure before submiting the code that you're strictly following this instruction.";
                md +=
                    "4. **Loop Closure:** For loops end with `Next`, Do While with `Loop`, While with `Wend`\n\n";

                md += "### Variable & Object Management\n";
                md +=
                    "1. **Declaration:** Separate variable declaration from initialization\n";
                md += "   ```\n";
                md += "   Dim myVar\n";
                md += '   myVar = "value"\n';
                md += "   ```\n";
                md +=
                    "2. **Object Assignment:** Always use `Set` keyword for objects\n";
                md += "   ```\n";
                md += '   Set rs = Server.CreateObject("ADODB.Recordset")\n';
                md += "   Set rs = Nothing\n";
                md += "   ```\n";
                md +=
                    "3. **Variant Type:** All variables are variants; no explicit typing allowed\n\n";

                md += "### Method & Function Invocation\n";
                md +=
                    "1. **Subs/Methods (no return):** Either omit parentheses or use Call keyword\n";
                md += "   ```\n";
                md += '   Response.Write "Hello"\n';
                md += '   Call Response.Write("Hello")\n';
                md += "   ```\n";
                md +=
                    "2. **Functions (with return):** Always use parentheses\n";
                md += "   ```\n";
                md += '   result = Len("text")\n';
                md += "   ```\n\n";

                md += "### Control Flow & Evaluation\n";
                md +=
                    "1. **Short-Circuit Logic:** VBScript does NOT short-circuit; evaluate both operands\n";
                md +=
                    "2. **Safe Array Access:** Nest conditions instead of using And\n";
                md += "   ```\n";
                md += "   If UBound(arr) >= 1 Then\n";
                md += "       If arr(1) = value Then\n";
                md += "           ' safe\n";
                md += "       End If\n";
                md += "   End If\n";
                md += "   ```\n";
                md += "3. **Error Handling:**\n";
                md += "   ```\n";
                md += "   On Error Resume Next\n";
                md += "   ' code\n";
                md += "   If Err.Number <> 0 Then\n";
                md += "       ' handle\n";
                md += "   End If\n";
                md += "   On Error GoTo 0\n";
                md += "   ```\n\n";

                md += "### String Operations\n";
                md +=
                    "1. **Concatenation:** Always use `&` operator, never `+`\n";
                md +=
                    "2. **Comparison:** Single `=` for equality (also used for assignment)\n";
                md +=
                    "3. **Line Continuation:** Use space followed by underscore `_` to break long lines\n\n";

                md += "### Server Object Creation\n";
                md +=
                    "1. **ProgID Format:** Use exact ProgID strings as documented\n";
                md +=
                    "2. **AxonASP Libraries:** Prefer native functions for maximum performance\n";
                md +=
                    "3. **Cleanup:** Always release objects when finished\n\n";

                md += "## Server-Side Objects Available\n\n";
                md +=
                    "AxonASP natively supports these custom libraries (preferred over custom code):\n";
                md += "- **G3JSON:** JSON encoding/decoding\n";
                md += "- **G3DB:** Database operations\n";
                md += "- **G3Files:** File system operations\n";
                md += "- **G3HTTP:** HTTP client requests\n";
                md += "- **G3MAIL:** Email sending\n";
                md += "- **G3Image:** Image processing\n";
                md += "- **G3Template:** Template rendering\n";
                md += "- **G3PDF:** PDF generation\n";
                md += "- **G3ZIP:** ZIP compression\n";
                md += "- **G3Crypto:** Cryptographic operations\n";
                md += "## Critical VBScript Pitfalls to Avoid\n\n";
                md += "### Subscript Out of Range\n";
                md +=
                    "**Problem:** Evaluating array indices that do not exist\n";
                md +=
                    "**Cause:** VBScript evaluates all And operands even if first is false\n";
                md += "**Solution:** Use nested If blocks for array safety\n\n";

                md += "### String vs. Numeric Addition\n";
                md +=
                    "**Problem:** Using `+` for concatenation causes implicit type coercion\n";
                md +=
                    "**Solution:** Always use `&` for string concatenation\n\n";

                md += "### Object Memory Leaks\n";
                md += "**Problem:** Forgetting to set objects to Nothing\n";
                md +=
                    "**Solution:** Always explicitly release with `Set objVar = Nothing`\n\n";

                md += "### Comparison Pitfalls\n";
                md +=
                    "**Problem:** Empty string vs. Null vs. Nothing are different\n";
                md +=
                    "**Solution:** Use IsEmpty(), IsNull(), IsNothing() for proper checks\n\n";

                md += "## Development Workflow\n\n";
                md += "1. Understand the business requirements above\n";
                md += "2. Design database schema with normalized tables\n";
                md += "3. Create models for data access and validation\n";
                md += "4. Build controllers to handle routing and logic\n";
                md +=
                    "5. Implement views for user interface, if needed with template support\n";
                md += "6. Test all code paths with sample data\n";
                md += "7. Verify error handling and edge cases\n";
                md += "8. Review for performance and memory efficiency\n\n";

                md += "\n\n";
                md += "## VBScript Classic ASP Compliance Summary\n\n";
                md +=
                    "The following guidelines from Microsoft Classic ASP standards must be enforced throughout development:\n\n";

                md += "### 1. ASP Directives and Delimiters\n";
                md +=
                    "- Page directive must exist entirely on the first line only\n";
                md += "- Keep directives compact with no extra spacing\n";
                md += "- Never leave tags unclosed or mismatched\n\n";

                md += "### 2. Control Structures (The If-Then Rule)\n";
                md += "- Never use single-line If statements\n";
                md += "- Always use block syntax with explicit End If\n";
                md += "- After Then always start a new line\n";
                md += "- Close loops with Next, Loop, or Wend\n\n";

                md += "### 3. Variable Declaration & Initialization\n";
                md += "- Enforce Option Explicit on every page\n";
                md += "- Declare and assign in separate statements\n";
                md += "- No typed variables (VBScript uses variants only)\n\n";

                md += "### 4. Object Assignment (Set vs. Let)\n";
                md += "- Always use Set for object references\n";
                md += "- Release objects with Set objVar = Nothing\n";
                md +=
                    "- Example: Set rs = Server.CreateObject(ADODB.Recordset)\n\n";

                md +=
                    "### 5. Method and Function Calling (Parentheses Rules)\n";
                md +=
                    "- Subs/methods without return: omit parentheses OR use Call keyword\n";
                md += "- Functions with return: always use parentheses\n";
                md +=
                    "- Example: Response.Write text or Call Response.Write(text)\n\n";

                md += "### 6. Major Quirks vs. Modern Languages\n";
                md +=
                    "- No short-circuit logic: both operands of AND/OR are evaluated\n";
                md += "- Use nested If blocks instead of compound conditions\n";
                md +=
                    "- Error handling: On Error Resume Next, check Err.Number, then On Error GoTo 0\n";
                md += "- String concatenation: always use &, never +\n";
                md += "- Line continuation: space followed by underscore _\n\n";

                md += "### 7. Server-Side Object Creation\n";
                md +=
                    "- Use exact ProgID strings documented for each library\n";
                md += "- Verify library is properly registered on server\n";
                md += "- Always check for creation errors\n\n";
                md +=
                    "- Avoid writing custom ASP code for functionality already available natively\nFollow these rules strictly\n";
                md += "\n\n---\n## Final Verification Checklist\n\n";
                md += "- [ ] All variables declared with Option Explicit\n";
                md +=
                    "- [ ] Objects assigned with Set and released with Nothing\n";
                md += "- [ ] String concatenation uses & operator\n";
                md += "- [ ] Array access protected with bounds checking\n";
                md += "- [ ] Error handling implemented throughout\n";
                md +=
                    "- [ ] No single-line/Inline If statements in code - this is critical, if fail the code will break \n";
                md += "- [ ] All loops properly closed (Next, Loop, Wend)\n";
                md += "- [ ] Method calls use correct parenthesis rules\n";
                md += "- [ ] Database connections properly closed\n";
                md += "- [ ] No hardcoded SQL values (use parameters)\n";
                md += "- [ ] File handles properly closed\n";
                md += "- [ ] Code tested and following Classic ASP pattern\n\n";

                document.getElementById("output-textarea").value = md;
                document.getElementById("output-area").classList.add("show");
                document
                    .getElementById("output-area")
                    .scrollIntoView({ behavior: "smooth", block: "nearest" });
            }

            function copyToClipboard() {
                var textarea = document.getElementById("output-textarea");
                if (!textarea.value) {
                    alert("Please generate document first.");
                    return;
                }

                textarea.select();
                try {
                    document.execCommand("copy");
                    var toast = document.getElementById("copy-toast");
                    toast.classList.add("show");
                    setTimeout(function () {
                        toast.classList.remove("show");
                    }, 2500);
                } catch (err) {
                    alert("Failed to copy to clipboard.");
                }
            }

            function clearForm() {
                if (confirm("Reset all fields to defaults?")) {
                    document.getElementById("description").value = "";
                    document.getElementById("appname").value = "";
                    document.getElementById("extra").value = "";
                    document.getElementById("style-mvc").checked = true;
                    document.getElementById("db-sqlite").checked = true;
                    document.getElementById("ui-axonasp").checked = true;
                    document.getElementById("color-light").checked = true;
                    document.getElementById("emoji-enabled").checked = true;
                    document.getElementById("feat-auth").checked = true;
                    document.getElementById("feat-sample").checked = true;
                    document.getElementById("feat-crud").checked = true;
                    document.getElementById("multilingual").checked = false;
                    document
                        .getElementById("output-area")
                        .classList.remove("show");
                    document
                        .querySelectorAll(
                            ".radio-item.selected, .check-item.selected"
                        )
                        .forEach(function (el) {
                            el.classList.remove("selected");
                        });
                    initMultiLang();
                }
            }

            // Initialize on load
            document.addEventListener("DOMContentLoaded", function () {
                initMultiLang();
            });
        </script>
    </body>
</html>
