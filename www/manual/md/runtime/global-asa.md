# global.asa

## Overview

`global.asa` is the application-level script file that defines Application and Session lifecycle event handlers and declares static objects for the `Application` and `Session` scopes. AxonASP compiles `global.asa` once at startup and runs the appropriate handlers when application or session events fire.

## File Location

Place `global.asa` at the root of the configured web root directory, alongside your default `default.asp` or `index.asp` file. Both the HTTP server and the FastCGI server look for the file in the configured web root.

When using the CLI (`axonasp-cli.exe`), `global.asa` must be placed in the **same directory as the CLI executable**. The CLI does not use the web root path for this file.

## Supported Event Handlers

Declare standard Sub procedures using the following exact names inside a `<SCRIPT LANGUAGE="VBScript" RUNAT="Server">` block:

| Event | Fired When |
|-------|------------|
| Application_OnStart | The application starts on the first request after server startup |
| Application_OnEnd | The application ends on clean server shutdown |
| Session_OnStart | A new session is created for a client |
| Session_OnEnd | A session expires or is abandoned |

```asp
<SCRIPT LANGUAGE="VBScript" RUNAT="Server">

Sub Application_OnStart()
    Application("SiteTitle") = "My AxonASP Site"
    Application("StartTime") = Now()
End Sub

Sub Application_OnEnd()
    ' Clean up application-level resources
End Sub

Sub Session_OnStart()
    Session("CreatedAt") = Now()
    Session("RequestCount") = 0
End Sub

Sub Session_OnEnd()
    ' Clean up session-level resources
End Sub

</SCRIPT>
```

## Static Object Declarations

Declare objects with Application or Session scope using the `<OBJECT>` tag. Static objects are automatically available throughout the application without calling `Server.CreateObject` on each page.

```asp
<OBJECT RUNAT="Server" SCOPE="Application" ID="AppCache" PROGID="G3JSON">
</OBJECT>

<OBJECT RUNAT="Server" SCOPE="Session" ID="UserData" PROGID="G3JSON">
</OBJECT>
```

**OBJECT tag attributes:**

| Attribute | Values | Description |
|-----------|--------|-------------|
| RUNAT | Server | Required. Must always be Server |
| SCOPE | Application / Session | Lifetime and sharing scope of the object |
| ID | identifier | Variable name used to access the object in ASP pages |
| PROGID | ProgID string | Object class identifier (same as the string passed to Server.CreateObject) |
| CLASSID | CLSID string | Alternative to PROGID for COM class identifiers |

## What Is Not Supported

The following Classic ASP `global.asa` features are not supported in AxonASP:

- **TypeLib declarations** — `<METADATA TYPE="TypeLib" ...>` directives are ignored.
- **Request and Response access in Application events** — `Request` and `Response` objects are not available inside `Application_OnStart` and `Application_OnEnd`. The server suppresses all response output from `global.asa` handlers to match IIS behavior.
- **#include directives** — File includes are not processed inside `global.asa`.
- **ObjectContext** — The `ObjectContext` intrinsic object is not available in event handlers.
- **On Error Resume Next suppressing compile errors** — A compile error in `global.asa` aborts server startup.

## Remarks

- `global.asa` is compiled once when the server starts. Restart the server to pick up changes to the file.
- UTF-8 BOM is automatically stripped from `global.asa` before compilation.
- Application and Session scope static objects declared in `global.asa` are initialized lazily when first accessed by an ASP page.
- `Session_OnEnd` fires when a session expires based on the configured timeout, not necessarily when the user closes the browser.
- `Application_OnEnd` is only fired during a clean server shutdown. An OS-level process kill will not trigger this event.
- The CLI uses the `global.asa` file located in the same directory as `axonasp-cli.exe`, not from a web root path.
- All four event handlers are optional. You can define only the events your application needs.

## Code Example

A typical `global.asa` that initializes a shared counter and tracks session creation time:

```asp
<SCRIPT LANGUAGE="VBScript" RUNAT="Server">

Sub Application_OnStart()
    Application("TotalVisits") = 0
    Application("AppStarted") = Now()
End Sub

Sub Session_OnStart()
    Application.Lock
    Application("TotalVisits") = Application("TotalVisits") + 1
    Application.Unlock
    Session("SessionStart") = Now()
End Sub

Sub Session_OnEnd()
    ' Nothing to clean up
End Sub

Sub Application_OnEnd()
    ' Nothing to clean up
End Sub

</SCRIPT>
```
