# Use JavaScript (JScript) in AxonASP Pages

## Overview

AxonASP provides a high-performance JavaScript (JScript) execution engine that allows you to write server-side logic using full ECMAScript 5 (ES5) standards for JScript compatibility and also ECMAScript 6 (ES6) and onward for new code. This page covers how to use JavaScript (JScript), use ASP intrinsic objects, and leverage modern JavaScript features within your ASP applications.

### Why use JavaScript (JScript) in ASP?
- **Familiar Syntax**: Many developers are more familiar with JavaScript than VBScript, making it easier to write complex logic. This allows the user to write full HTML pages, and services using javascript only, with the full support of the AxonASP framework, and in a way that is easier and more memory efficient than using systems like NodeJS.
- **Rich Ecosystem**: Access to a wide range of JavaScript libraries and tools.
- **Performance**: The AxonASP JavaScript engine is optimized for server-side execution, providing better performance for certain workloads.
- **ASP Intrinsic Objects**: Seamless access to ASP intrinsic objects like `Request`, `Response`, `Session`, and `Application` allows you to build dynamic web applications with ease.
- **Modern Features**: Support for ES6 and most of ES7+ features allows you to write cleaner and more efficient code.

### Unlocking the Power of Server-Side JavaScript with AxonASP

Building web applications and APIs using native JavaScript on the server has never been more intuitive. AxonASP empowers developers to write server-side logic utilizing the exact same JavaScript syntax and paradigms they already master in the browser, seamlessly bridging the gap between front-end and back-end development.

One of the most profound advantages of AxonASP's JavaScript engine is its synchronous execution model. Unlike Node.js, which frequently forces developers into complex, deeply nested asynchronous patterns—often colloquially known as "callback hell" or an endless tree of async/await promises—AxonASP executes JavaScript sequentially and synchronously by default. This dramatically simplifies control flow, state management, and debugging. You write your script naturally, just as you would for a standard web page, and the engine compiles and executes it with blazing speed. This synchronous predictability reduces overhead, minimizes cognitive load, and results in cleaner, highly maintainable code for creating robust APIs, handling database transactions, and rendering dynamic websites.

## Syntax
To set JavaScript (JScript) as the default language for an entire page, use the language directive at the very first line of your file:

```asp
<%@ Language="javascript" %>
```

or

```asp
<%@ Language="jscript" %>
```

Alternatively, you can use JavaScript (JScript) within specific script blocks:

```html
<script runat="server" language="javascript">
    // JavaScript (JScript) code here
</script>
```

You can also execute JavaScript files directly from the command line using the AxonASP CLI. See the CLI example below for details.

## Parameters and Arguments
- **Language Directive** (Required for page-level): The value must be `"javascript"` or `"jscript"`.
- **runat="server"** (Required for script tags): Ensures the code executes on the server rather than the client browser.
- **ASP Intrinsic Objects**: Native access to **Request**, **Response**, **Server**, **Session**, **Application**, and **Err**. Note that in JavaScript (JScript), these object names and their members are **case-sensitive**.
- **CLI Mode Flag**: When running JavaScript from the command line, use the `-m javascript` flag to force the JavaScript engine mode. Files with `.js` and `.mjs` extensions are automatically recognized as JavaScript in the default mode.

## Return Values
The JavaScript (JScript) engine returns standard JavaScript values (String, Number, Boolean, Object, Array, null, undefined). When communicating with the AxonASP VM or VBScript components:
- JavaScript objects are automatically converted to their closest AxonASP **Value** equivalent.
- **undefined** and **null** map to **Empty** in the VM context.
- JavaScript arrays map to AxonASP arrays, and JavaScript objects map to AxonASP objects.

## Remarks
- **ECMAScript 5/6 Support**: AxonASP's JavaScript (JScript) engine supports all ES5 features and most ES6 features, including JSON support (`JSON.parse`, `JSON.stringify`), and standard Array methods (`map`, `filter`, `reduce`). Most features from later versions are also supported, refer to the documentation for specific details.
- **Case Sensitivity**: Unlike VBScript, JavaScript (JScript) is strictly case-sensitive. You must use `Response.Write`, not `response.write`.
- **Engine Architecture**: JavaScript (JScript) execution in AxonASP utilizes a sophisticated Abstract Syntax Tree (AST) parser and interpreter, providing optimized performance for complex logic.
- **Global Console**: The engine includes a built-in **console** object (`console.log`, `console.warn`, `console.error`) for server-side debugging and diagnostics. Output is directed to the system console or log files depending on your `axonasp.toml` configuration.
- **Interoperability**: You can mix VBScript and JavaScript (JScript) in the same application by using separate `<script runat="server">` blocks, though global variable sharing follows standard ASP scoping rules.
- **CLI Execution**: JavaScript files can be executed directly from the command line without a web server. This is ideal for batch processing, automation tasks, and testing. Use the `-r` flag to run a file and `-m javascript` to enforce JavaScript mode.

## Code Examples

### Example 1: Building a JSON REST API Endpoint

The following example demonstrates creating a full JSON REST API endpoint using JavaScript. It parses query string parameters, processes data, and returns a structured JSON response with appropriate HTTP headers.

```asp
<%@ Language="javascript" %>
<% 
Response.ContentType = "application/json";

// Parse the HTTP method from the request
var method = Request.ServerVariables("REQUEST_METHOD");

// Define a simple in-memory data store
var items = [
    { id: 1, name: "Widget A", price: 9.99 },
    { id: 2, name: "Widget B", price: 14.99 },
    { id: 3, name: "Widget C", price: 19.99 }
];

if (method === "GET") {
    // Check if a specific ID was requested
    var idParam = Request.QueryString("id");
    if (idParam !== "") {
        var id = parseInt(idParam, 10);
        var found = null;
        for (var i = 0; i < items.length; i++) {
            if (items[i].id === id) {
                found = items[i];
                break;
            }
        }
        if (found) {
            Response.Write(JSON.stringify({ success: true, data: found }));
        } else {
            Response.Status = "404 Not Found";
            Response.Write(JSON.stringify({ success: false, error: "Item not found" }));
        }
    } else {
        // Return all items
        Response.Write(JSON.stringify({ success: true, count: items.length, data: items }));
    }
} else {
    Response.Status = "405 Method Not Allowed";
    Response.Write(JSON.stringify({ success: false, error: "Method not supported" }));
}

console.log("API request processed: " + method + " at " + new Date().toISOString());
%>
```

### Example 2: Generating a Dynamic HTML Web Page

This example shows how to build a complete dynamic HTML page using JavaScript, demonstrating string templating, data iteration, and conditional rendering within a server-side script.

```asp
<%@ Language="javascript" %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Product Catalog</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        .product { border: 1px solid #ddd; padding: 15px; margin-bottom: 10px; border-radius: 5px; }
        .product h3 { margin: 0 0 5px 0; }
        .price { color: green; font-weight: bold; }
    </style>
</head>
<body>
    <h1>Product Catalog</h1>
    <p>Generated on: <%= new Date().toString() %></p>
<%
var products = [
    { name: "Wireless Mouse", price: 29.99, inStock: true },
    { name: "Mechanical Keyboard", price: 89.99, inStock: true },
    { name: "USB-C Hub", price: 49.99, inStock: false },
    { name: "Monitor Stand", price: 39.99, inStock: true }
];

for (var i = 0; i < products.length; i++) {
    var p = products[i];
%>
    <div class="product">
        <h3><%= p.name %></h3>
        <p class="price">$<%= p.price.toFixed(2) %></p>
        <p>Status: <%= p.inStock ? "In Stock" : "Out of Stock" %></p>
    </div>
<%
}
%>
    <p>Total products displayed: <%= products.length %></p>
</body>
</html>
```

### Example 3: global.asa Using JavaScript

The `global.asa` file can also be written using JavaScript. This example shows how to define session-level and application-level event handlers using the JavaScript language directive.

```asp
<script language="javascript" runat="server">
function Application_OnStart() {
    Application.Lock();
    Application("appName") = "AxonASP Demo";
    Application("startTime") = new Date().toISOString();
    Application("visitorCount") = 0;
    Application.Unlock();
    console.log("Application started: " + Application("startTime"));
}

function Session_OnStart() {
    Application.Lock();
    Application("visitorCount") = Application("visitorCount") + 1;
    var count = Application("visitorCount");
    Application.Unlock();
    Session("sessionId") = "SESS" + Math.floor(Math.random() * 1000000);
    Session("startTime") = new Date().toISOString();
    console.log("Session started. Total visitors: " + count);
}

function Session_OnEnd() {
    console.log("Session ended: " + Session("sessionId"));
}

function Application_OnEnd() {
    console.log("Application shutting down. Started at: " + Application("startTime"));
}
</script>
```

### Example 4: Running JavaScript from the CLI (Writing to a File)

The AxonASP CLI allows you to execute JavaScript files directly from the command line without a web server. This is useful for batch processing, data transformation, scheduled tasks, and system automation.

Create a file named `generate-report.js` with the following content:

```javascript
<%@ Language="javascript" %>
<%
// Generate a simple HTML report and write it to disk
var G3FILES = Server.CreateObject("G3FILES");

var reportData = [
    { task: "System Health Check", status: "Passed", timestamp: new Date().toISOString() },
    { task: "Database Connectivity", status: "Passed", timestamp: new Date().toISOString() },
    { task: "Disk Space Verification", status: "Warning", timestamp: new Date().toISOString() }
];

var html = "<!DOCTYPE html>\n";
html += "<html><head><title>AxonASP CLI Report</title></head><body>\n";
html += "<h1>Automated Report</h1>\n";
html += "<table border='1'><tr><th>Task</th><th>Status</th><th>Timestamp</th></tr>\n";

for (var i = 0; i < reportData.length; i++) {
    var row = reportData[i];
    html += "<tr><td>" + row.task + "</td><td>" + row.status + "</td><td>" + row.timestamp + "</td></tr>\n";
}
html += "</table>\n";
html += "<p>Generated by AxonASP CLI</p>\n";
html += "</body></html>";

// Write the report to a file using G3FILES
G3FILES.Write("report-output.html", html);
Response.Write("Report generated successfully: report-output.html\n");
console.log("CLI report written to report-output.html");
%>
```

Then execute it from your terminal:

```powershell
> axonasp-cli.exe -r generate-report.js -m javascript
```

The `-r` flag tells the CLI to run the specified file directly and exit. The `-m javascript` flag ensures the JavaScript engine mode is used. The script has full access to ASP intrinsic objects (`Response`, `Server`, `Request`), custom libraries (`G3FILES`, `G3JSON`, etc.), and the `console` object for logging. Output is written to stdout and can be redirected to a file if needed.

You can also run JavaScript files with the `.asp` extension in default mode:

```powershell
> axonasp-cli.exe -r myscript.asp
```
