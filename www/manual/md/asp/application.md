# Use the ASP Application Object

## Overview
The **Application** object in G3Pix AxonASP stores data that is shared across all users and all requests for the entire lifetime of the web application. Unlike the **Session** object, which is unique to each visitor, **Application** values are global. This object is typically used for site-wide settings, application-level counters, or caching shared lookup data.

## Prerequisites
- **Global.asa**: Initial application values are often set in the `Application_OnStart` event.
- **Concurrency Awareness**: Because multiple requests access this object simultaneously, you must use locking methods to prevent data collisions.

## How it Works
The **Application** object acts as a global dictionary. When you store a value, it remains in memory until the AxonASP server process is restarted.
- **Locking**: Use the **Lock** method before modifying any value to ensure thread safety.
- **Unlocking**: Always call the **Unlock** method immediately after your modification is complete.
- **Persistence**: Data is lost if the server shuts down or the application pool is recycled.

## API Reference

### Methods
- **Lock**: Prevents other clients from modifying the **Application** object. Returns **Empty**.
- **Unlock**: Releases the lock on the **Application** object, allowing other clients to modify it. Returns **Empty**.

### Collections
- **Contents**: Returns a collection of all values added to the application through script commands.
- **StaticObjects**: Returns a collection of all objects added to the application with the `<OBJECT>` tag in `global.asa`.

## Code Example
The following example demonstrates how to safely increment a global page hit counter using the **Lock** and **Unlock** methods.

```asp
<%
' Access the Application object
Application.Lock
Application("TotalHits") = Application("TotalHits") + 1
Application.Unlock

Response.Write "Total application views: " & Application("TotalHits")
%>
```

## Remarks
- **Thread Safety**: Failing to use **Lock** and **Unlock** when writing to the **Application** object can result in lost updates in high-concurrency environments.
- **Memory Management**: Avoid storing large objects or frequently changing data in the **Application** object to prevent memory bloat.
- **Redirection**: Never call `Response.Redirect` or `Response.End` while an application lock is active, as this may bypass the **Unlock** call and freeze the object for other threads.
