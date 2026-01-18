<%@ Language=VBScript %>
<html>
<head>
    <title>G3pix AxonASP Interpreter - Test Suite</title>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            overflow-x: hidden;
        }
        .main-wrapper {
            display: flex;
            height: 100vh;
        }
        .sidebar {
            width: 50%;
            overflow-y: auto;
            padding: 30px;
            background: #f9f9f9;
            border-right: 2px solid #ddd;
        }
        .content-area {
            width: 50%;
            display: flex;
            flex-direction: column;
            background: #fff;
        }
        .iframe-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #fff;
            padding: 15px 20px;
            font-weight: bold;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .close-btn {
            background: rgba(255,255,255,0.2);
            border: none;
            color: #fff;
            padding: 5px 10px;
            cursor: pointer;
            border-radius: 4px;
            transition: background 0.2s;
        }
        .close-btn:hover {
            background: rgba(255,255,255,0.3);
        }
        #test-iframe {
            flex: 1;
            border: none;
            width: 100%;
            height: 100%;
        }
        .no-test-selected {
            display: flex;
            align-items: center;
            justify-content: center;
            height: 100%;
            color: #999;
            font-size: 18px;
            text-align: center;
            background: #fafafa;
        }
        h1 {
            color: #333;
            margin-bottom: 20px;
            font-size: 1.8em;
        }
        .category {
            margin-bottom: 30px;
        }
        .category-title {
            color: #764ba2;
            font-weight: bold;
            font-size: 1.1em;
            margin-bottom: 12px;
            padding-bottom: 8px;
            border-bottom: 2px solid #667eea;
        }
        .test-item {
            margin-bottom: 10px;
            padding: 12px;
            background: #fff;
            border-radius: 4px;
            border: 1px solid #ddd;
            transition: all 0.2s;
            cursor: pointer;
        }
        .test-item:hover {
            border-color: #667eea;
            box-shadow: 0 2px 8px rgba(102, 126, 234, 0.2);
            transform: translateX(4px);
        }
        .test-item.active {
            background: #667eea;
            color: #fff;
            border-color: #667eea;
        }
        .test-item.active .test-title,
        .test-item.active .test-desc {
            color: #fff;
        }
        .test-title {
            font-weight: bold;
            color: #333;
            font-size: 0.95em;
            margin-bottom: 5px;
        }
        .test-desc {
            font-size: 0.8em;
            color: #666;
            line-height: 1.4;
        }
        a.test-link {
            color: inherit;
            text-decoration: none;
            display: block;
        }
        a.test-link:hover {
            text-decoration: none;
        }
        hr {
            margin: 20px 0;
            border: none;
            border-top: 1px solid #ddd;
        }
        .footer {
            text-align: center;
            color: #999;
            font-size: 0.8em;
            padding-top: 15px;
            margin-top: 20px;
            border-top: 1px solid #ddd;
        }
    </style>
</head>
<body>
    <div class="main-wrapper">
        <!-- Sidebar with test links -->
        <div class="sidebar">
            <h1>Tests & Tools</h1>
            
            <!-- Core Functionality -->
            <div class="category">
                <div class="category-title">Core Functionality</div>
                
                <div class="test-item" onclick="loadTest('test_basics.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Basic Syntax & Logic</div>
                        <div class="test-desc">Variables, If/Else, Loops, Subroutines, Session, Arrays.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_flow.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Control Flow</div>
                        <div class="test-desc">For/Next with Step, Do/Loop, Select/Case, If/ElseIf/Else.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_directive.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">ASP Directives</div>
                        <div class="test-desc">Tests <%@ Language=VBScript %> directive support.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_operators.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Operators</div>
                        <div class="test-desc">Mathematical and logical operators: Mod, \, And, Or, Not, +, -, *, /, Null, Empty, Nothing, True, False.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_select_case_multi.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Select Case (multi)</div>
                        <div class="test-desc">Multiple values and numeric ranges in Case clauses.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_with.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">With Statement</div>
                        <div class="test-desc">Implicit object context for properties and methods.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_operators_math_logic.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Math & Logic Operators</div>
                        <div class="test-desc">Mod, \, And, Or, Not operators (bitwise and logical).</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_type_functions.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Type Functions & Literals</div>
                        <div class="test-desc">TypeName, VarType, RGB, IsObject, IsEmpty, IsNull, Is Nothing, hex/octal literals.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_const_datetime.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Const Declarations & DateTime</div>
                        <div class="test-desc">Const read-only declarations, DateAdd, DateDiff, DatePart, FormatDateTime, and more.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_hex_octal.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Hexadecimal & Octal Literals</div>
                        <div class="test-desc">VBScript &amp;h (hex) and &amp;o (octal) numeric literals.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_features.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Features & Operators</div>
                        <div class="test-desc">Math operators (*, -, /), Subroutine parameters, string operations.</div>
                    </a>
                </div>
            </div>

            <!-- Built-in Functions -->
            <div class="category">
                <div class="category-title">Built-in Functions</div>
                
                <div class="test-item" onclick="loadTest('test_functions.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Standard Functions</div>
                        <div class="test-desc">String, Date, Math, Type conversion functions.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_keywords.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Keywords & Types</div>
                        <div class="test-desc">Empty, Null, Nothing, True, False, type checking functions.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_const.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Constants</div>
                        <div class="test-desc">Const declarations and reassignment protection.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_other_functions.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Other Functions</div>
                        <div class="test-desc">ScriptEngine, TypeName, VarType, RGB, IsObject, Eval.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_datefix.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Date Functions</div>
                        <div class="test-desc">DateDiff, DatePart, DateAdd, DateValue, DateSerial, FormatDateTime.</div>
                    </a>
                </div>
            </div>

            <!-- Server & Objects -->
            <div class="category">
                <div class="category-title">Server & Objects</div>
                
                <div class="test-item" onclick="loadTest('test_objects.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Core Objects</div>
                        <div class="test-desc">Request, Response, Server, Session objects.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_application.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Application Object</div>
                        <div class="test-desc">Lock, Unlock, StaticObjects, variable enumeration.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_include.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Server-Side Includes</div>
                        <div class="test-desc">SSI directives (file and virtual) with recursive processing.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_session_demo.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Session File Storage</div>
                        <div class="test-desc">Session persistence in temp/session/ directory with JSON files.</div>
                    </a>
                </div>
            </div>

            <!-- Advanced Features -->
            <div class="category">
                <div class="category-title">Advanced Features</div>
                
                <div class="test-item" onclick="loadTest('test_error.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Error Handling</div>
                        <div class="test-desc">Runtime errors and ASPError object (requires debug mode).</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_debug.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Debug Tools</div>
                        <div class="test-desc">Debugging utilities and variable inspection.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_redirect.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Redirects</div>
                        <div class="test-desc">HTTP redirection and Response.Redirect.</div>
                    </a>
                </div>
            </div>

            <!-- New Features (v2) -->
            <div class="category">
                <div class="category-title">New Features (v2)</div>
                
                <div class="test-item" onclick="loadTest('test_global_events.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Global.asa & Events</div>
                        <div class="test-desc">Application_OnStart, Session_OnStart, and global variables.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_response_new.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Response Object V2</div>
                        <div class="test-desc">Headers, Caching, Charset, BinaryWrite, Logging.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_server_advanced.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Server Object V2</div>
                        <div class="test-desc">Server.Execute, Server.Transfer, ScriptTimeout.</div>
                    </a>
                </div>
            </div>

            <!-- Interpreter Upgrades -->
            <div class="category">
                <div class="category-title">Interpreter Upgrades</div>
                
                <div class="test-item" onclick="loadTest('test_arrays_redim.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Array Redim & Preserve</div>
                        <div class="test-desc">Dynamic arrays and ReDim Preserve.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_error_handling.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Error Handling (Try/Catch)</div>
                        <div class="test-desc">On Error Resume Next and Err Object.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_components.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">COM Components</div>
                        <div class="test-desc">Server.CreateObject (Dictionary, XMLHTTP).</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_regexp.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">VBScript RegExp</div>
                        <div class="test-desc">Regular Expressions (Pattern, Execute, Replace).</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_regexp_new.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">New RegExp Syntax</div>
                        <div class="test-desc">Testing 'Set x = New RegExp'.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_regexp_property.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">RegExp Property Assignment</div>
                        <div class="test-desc">Testing object property assignment (objRegExp.IgnoreCase = True).</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_debug_server.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Debug Server Object</div>
                        <div class="test-desc">Deep inspection of CreateObject flow.</div>
                    </a>
                </div>
            </div>

            <!-- Language Upgrades (v3) -->
            <div class="category">
                <div class="category-title">Language Upgrades (v3)</div>
                
                <div class="test-item" onclick="loadTest('test_features_new.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Date Literals</div>
                        <div class="test-desc">Parsing #date# literals and date math.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_option_explicit_pass.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Option Explicit (Pass)</div>
                        <div class="test-desc">Tests explicit variable declaration (Success case).</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_option_explicit_fail.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Option Explicit (Fail)</div>
                        <div class="test-desc">Tests failure on undeclared variables (Expect Error).</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_classes.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Classes</div>
                        <div class="test-desc">Class Definition, Properties, Methods, and Scope.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_math.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Math Operators</div>
                        <div class="test-desc">Exponentiation (^) operator and math expressions.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_response_extras.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Response Extras</div>
                        <div class="test-desc">PICS, CacheControl interaction, and Client Connection.</div>
                    </a>
                </div>
            </div>

            <!-- Additional Tests -->
            <div class="category">
                <div class="category-title">Additional Tests</div>
                
                <div class="test-item" onclick="loadTest('test_crypto.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Cryptography & UUID</div>
                        <div class="test-desc">UUID generation, BCrypt password hashing and verification.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_file.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">File Operations</div>
                        <div class="test-desc">File system functions: read, write, delete and list.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_fso.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">FSO Legacy</div>
                        <div class="test-desc">Legacy FileSystemObject compatibility and File.* helper demo.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_adodb.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">ADODB Database Access</div>
                        <div class="test-desc">ADODB.Connection and ADODB.Recordset for SQLite, MySQL, PostgreSQL and MS SQL.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_adodb_advanced.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">ADODB Advanced Features</div>
                        <div class="test-desc">Recordset.Supports(), Filter property, Connection.Errors, and ADODB.Stream.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_adodb_comprehensive.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">ADODB Comprehensive Test</div>
                        <div class="test-desc">Complete ADODB feature testing and validation.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_adodb_stream_final.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">ADODB.Stream Implementation</div>
                        <div class="test-desc">ADODB.Stream object for advanced file operations.</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_file_nil_check.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">File Nil/Empty Path Validation</div>
                        <div class="test-desc">Tests proper handling of nil/empty paths in ADODB.Stream and G3FILES (Security Fix).</div>
                    </a>
                </div>
                
                <div class="test-item" onclick="loadTest('test_json.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">JSON Operations</div>
                        <div class="test-desc">JSON object creation, serialization, parsing and iteration.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_mail.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Mail Library</div>
                        <div class="test-desc">Email sending capabilities (SMTP &amp; Environment variables).</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_template.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Template Engine</div>
                        <div class="test-desc">Go Template rendering with ASP data.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_modern.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Modern Features</div>
                        <div class="test-desc">Modern ASP features like environment variables and API consumption.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_msxml_simple.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">MSXML Simple Tests</div>
                        <div class="test-desc">MSXML2.ServerXMLHTTP and MSXML2.DOMDocument simple operations.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_msxml.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">MSXML Objects (XML)</div>
                        <div class="test-desc">MSXML2.ServerXMLHTTP and MSXML2.DOMDocument for XML processing and HTTP.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_msxml_full.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">MSXML Complete Suite</div>
                        <div class="test-desc">Complete MSXML2 functionality testing (Fixed).</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_dim_colon_and_local_vars.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Dim with Colon & Local Variables</div>
                        <div class="test-desc">Dim var : var = value syntax and local variables in functions.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_user_pre_function.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">User Function 'pre' Test</div>
                        <div class="test-desc">User-defined functions with local Dim and proper variable scoping.</div>
                    </a>
                </div>
            </div>

            <!-- Experimental & Debug Tests -->
            <div class="category">
                <div class="category-title">Response & Operators Tests</div>

                <div class="test-item" onclick="loadTest('test_response_contenttype.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Response.ContentType Implementation</div>
                        <div class="test-desc">Response.ContentType property now returns as valid HTTP header.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_operators_math_logic.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Mathematical & Logical Operators</div>
                        <div class="test-desc">Complete test suite for Mod, \, And, Or, Not, and bitwise operators.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_content_type.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Content-Type JSON Test</div>
                        <div class="test-desc">Test setting Response.ContentType to application/json.</div>
                    </a>
                </div>
            </div>

            <!-- Experimental & Debug Tests -->
            <div class="category">
                <div class="category-title">Experimental & Debug Tests</div>

                <div class="test-item" onclick="loadTest('test_asplite_fixes.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">ASPLite Fixes</div>
                        <div class="test-desc">Critical fixes: AscW/ChrW, Dictionary enumeration, VarType.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_arrays_redim.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Arrays & ReDim</div>
                        <div class="test-desc">Dynamic array creation with ReDim Preserve and VarType.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_basic_ok.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Basic OK Test</div>
                        <div class="test-desc">Simple sanity check for ASP interpreter functionality.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_byref_simple.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">ByRef Parameter Test</div>
                        <div class="test-desc">Function parameters passed by reference using ByRef.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_call_simple.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Method Call Test</div>
                        <div class="test-desc">Calling object methods with parentheses notation.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_debug_colon.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Debug Colon</div>
                        <div class="test-desc">Testing colon syntax in variable declarations.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_debug_pre_function.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Function 'pre' Debug</div>
                        <div class="test-desc">Testing different calling styles for user functions.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_dim_colon_debug.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Dim with Colon Debug</div>
                        <div class="test-desc">Testing Dim variable declaration with colon syntax.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_direct_stream.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Direct ADODB.Stream</div>
                        <div class="test-desc">Direct ADODB.Stream testing and validation.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_ebook_replica.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">E-Book Replica Test</div>
                        <div class="test-desc">Testing specific e-book.asp replication case.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_exact_user_case.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">User Case Test</div>
                        <div class="test-desc">Tests specific user-reported case with functions.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_executeglobal_simple.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">ExecuteGlobal Test</div>
                        <div class="test-desc">Testing ExecuteGlobal dynamic code execution.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_exglobal_minimal.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">ExecuteGlobal Minimal</div>
                        <div class="test-desc">Minimal ExecuteGlobal functionality test.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_func_simple.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Function Simple Test</div>
                        <div class="test-desc">Basic function definition and calling test.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_function_simple2.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Function Simple Test 2</div>
                        <div class="test-desc">Second basic function definition test.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_global_events.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Global Events Test</div>
                        <div class="test-desc">Application_OnStart, Session_OnStart events.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_hello.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Hello World Test</div>
                        <div class="test-desc">Basic Response.Write functionality test.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_isnot.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Is Not & WriteSafe</div>
                        <div class="test-desc">Is Not operator and Document.WriteSafe methods.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_load_return.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Load & Return Test</div>
                        <div class="test-desc">Testing load and return statement behavior.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_mappath.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">MapPath Test</div>
                        <div class="test-desc">Server.MapPath functionality testing.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_minimal.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Minimal Test</div>
                        <div class="test-desc">Minimal ASP functionality test.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_minimal_dup.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Minimal Duplicate Test</div>
                        <div class="test-desc">Minimal duplicate functionality test.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_mult_simple.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Multiplication Simple</div>
                        <div class="test-desc">Simple multiplication operator testing.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_option_explicit.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Option Explicit Test</div>
                        <div class="test-desc">Option Explicit variable declaration mode.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_operators_math_logic.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Math & Logic Operators</div>
                        <div class="test-desc">Mathematical and logical operators testing.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_pre_complex.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Pre Complex Args</div>
                        <div class="test-desc">Testing pre function with complex arguments.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_pre_option_explicit.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Pre with Option Explicit</div>
                        <div class="test-desc">Function pre() with Option Explicit mode.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_pre_tag.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Pre Tag Test</div>
                        <div class="test-desc">Testing pre HTML tag compatibility.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_public_func.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Public Function Test</div>
                        <div class="test-desc">Public function declaration and usage.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_regexp_class.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">RegExp Class Test</div>
                        <div class="test-desc">RegExp object class functionality.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_regexp_debug.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">RegExp Debug</div>
                        <div class="test-desc">RegExp object debugging and inspection.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_regexp_explicit.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">RegExp Explicit Test</div>
                        <div class="test-desc">RegExp with Option Explicit mode.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_regexp_simple.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">RegExp Simple Test</div>
                        <div class="test-desc">Simple RegExp functionality test.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_removecrb.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Remove CRB Test</div>
                        <div class="test-desc">Testing carriage return/line break removal.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_response_extras.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Response Extras</div>
                        <div class="test-desc">Response object advanced features.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_response_new.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Response Object V2</div>
                        <div class="test-desc">Headers, Caching, Charset, BinaryWrite.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_select_case_multi.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Select Case Multi</div>
                        <div class="test-desc">Multiple values and ranges in Case clauses.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_server_advanced.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Server Object V2</div>
                        <div class="test-desc">Server.Execute, Server.Transfer, ScriptTimeout.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_server_child.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Server Child Test</div>
                        <div class="test-desc">Child server execution and transfer.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_simple.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Simple Test</div>
                        <div class="test-desc">Minimal test of basic ASP functionality.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_simple_colon.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Simple Colon Test</div>
                        <div class="test-desc">Minimal Dim colon syntax test.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_simple_pre.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Simple Pre Function</div>
                        <div class="test-desc">Testing pre function in simple context.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_simple_start.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Simple Start Test</div>
                        <div class="test-desc">Simple startup test case.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_stream.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">ADODB.Stream Test</div>
                        <div class="test-desc">ADODB.Stream object functionality.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_stream_debug.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Stream Debug Test</div>
                        <div class="test-desc">ADODB.Stream debugging and inspection.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_stream_detail.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Detailed Stream Debug</div>
                        <div class="test-desc">Detailed ADODB.Stream debugging.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_stream_final.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Stream Final Test</div>
                        <div class="test-desc">Final user-defined stream() function test.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_stream_final_working.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Stream Final Working</div>
                        <div class="test-desc">Working user stream() function implementation.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_stream_minimal.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Simple Stream Test</div>
                        <div class="test-desc">Simple ADODB.Stream test.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_stream_new.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Stream Debug New</div>
                        <div class="test-desc">New stream debugging approach.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_stream_simple.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Stream Direct Test</div>
                        <div class="test-desc">Direct ADODB.Stream testing.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_stream_working.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">Stream Working Test</div>
                        <div class="test-desc">Working stream functionality test.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_user_function.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">User stream() Function</div>
                        <div class="test-desc">Testing user-defined stream() function.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_user_stream.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">User Stream Function</div>
                        <div class="test-desc">Testing user stream function implementation.</div>
                    </a>
                </div>

                <div class="test-item" onclick="loadTest('test_with.asp', this)">
                    <a href="javascript:void(0)" class="test-link">
                        <div class="test-title">With Statement Test</div>
                        <div class="test-desc">With statement for implicit object context.</div>
                    </a>
                </div>
            </div>

            <hr>
            <div class="footer">
                <strong>G3pix AxonASP</strong><br>
                Classic ASP Interpreter in Go<br>
                All content served with UTF-8 encoding
            </div>
        </div>

        <!-- Main content area with iframe -->
        <div class="content-area">
            <div class="iframe-header">
                <span id="current-test">Select a test from the left</span>
                <button class="close-btn" onclick="closeTest()">Close</button>
            </div>
            <div id="test-container" class="no-test-selected">
                Select a test to view it here
            </div>
        </div>
    </div>

    <script>
        function loadTest(testFile, element) {
            // Remove active class from all items
            document.querySelectorAll('.test-item').forEach(item => {
                item.classList.remove('active');
            });
            
            // Add active class to clicked item
            element.classList.add('active');
            
            // Update header
            const testTitle = element.querySelector('.test-title').textContent;
            document.getElementById('current-test').textContent = testTitle;
            
            // Create and show iframe
            const container = document.getElementById('test-container');
            container.innerHTML = '<iframe id="test-iframe" src="' + testFile + '"></iframe>';
        }

        function closeTest() {
            const container = document.getElementById('test-container');
            container.innerHTML = '<div class="no-test-selected">Select a test to view it here</div>';
            document.getElementById('current-test').textContent = 'Select a test from the left';
            document.querySelectorAll('.test-item').forEach(item => {
                item.classList.remove('active');
            });
        }
    </script>
</body>
</html>
