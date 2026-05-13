# G3AXONLIVE Object Overview

The `G3AXONLIVE` object is the native server-side procedural controller for the AxonLive framework. By utilizing this object, developers can receive asynchronous JavaScript `fetch` events from the frontend, process the associated business logic locally in Go memory, and respond with targeted DOM updates (HTML patches) and client instructions.

## ProgID

```vbscript
Set obj = Server.CreateObject("G3AXONLIVE")
```

```javascript
var obj = Server.CreateObject("G3AXONLIVE");
```

## Architecture

The `G3AXONLIVE` object acts as a bridge between the incoming async POST request (triggered by `g3axonlive.js`) and the backend ASP engine. It provides methods to identify the source of the event (`EventComponentID`), fetch contextual data (`GetEventArg`), and securely update state.

Async event routing is bound to the authenticated ASP session cookie (`ASPSESSIONID`). The server validates event session identity against the authenticated request context before resolving the target page.

When an async request is handled, the framework requires developers to call `EndAsyncResponse()`. This method collects all registered components and actions, serializes them into a JSON payload, writes it to the output buffer, and securely halts the ASP execution before the rest of the HTML template is rendered.

## State Management

G3AxonLive includes built-in granular persistence. The methods `SetComponentProperty` and `GetComponentProperty` bypass the standard ASP `Session` object and store data directly in an optimized, concurrent memory map keyed by the current Session ID. This guarantees memory efficiency and strict isolation across users.