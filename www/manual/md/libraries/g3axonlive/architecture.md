# Architecture and Request Lifecycle

G3AxonLive's architecture is designed for performance and security by centralizing core logic in the native Go backend. The ASP script's role is to handle business logic, not to manage the complexities of async communication.

## Core Components

1.  **`g3axonlive.js` (Client-Side Engine):** A lightweight vanilla JavaScript bridge that runs in the user's browser. It automatically intercepts events from designated components, sends them to the server via `fetch`, and applies the resulting DOM patches.
2.  **`/g3al/` Endpoint:** A dedicated HTTP endpoint within AxonASP that receives all async requests from the client-side engine. It identifies the user's session and the correct ASP page to execute.
3.  **`G3AXONLIVE` Go Object:** The native `Server.CreateObject("G3AXONLIVE")` instance. This object is the heart of the framework, providing the procedural API to interact with the request lifecycle, manage state, and build the response.
4.  **Go State Store:** A process-wide, in-memory store (a Go singleton) that securely holds all component and session data. This state is never exposed directly to the client.

---

## Request Lifecycle: Initial Page Load

This is the standard, full-page request that happens when a user first visits your `.asp` page.

```
+----------+      1. HTTP GET /mypage.asp      +-----------------+
|          | --------------------------------> |                 |
| Browser  |                                   |   AxonASP Server  |
|          |      2. Execute mypage.asp        | (Go Backend)    |
|          | <-------------------------------  |                 |
|          |   - AxonLive.InitPage()           |                 |
|          |     (registers session)           |                 |
|          |   - Render full HTML              |                 |
+----------+                                   +-----------------+
     |
     | 3. Full HTML response sent to browser
     |    - Includes <script src="/axonlive/g3axonlive.js">
     |    - Includes G3AxonLive.init("SESSION_ID")
     |    - Components have data-g3al-id attributes
     v
+----------+
|          |
|  Page    |
| Rendered |
+----------+
```

1.  The user's browser sends a standard GET request for your ASP page.
2.  AxonASP executes your script. The call to `AxonLive.InitPage()` during this initial load is crucial: it registers the mapping between the user's `Session.SessionID` and the URL of the current script (`mypage.asp`). This tells the `/g3al/` endpoint which script to run for future async updates.
3.  Your page renders the full initial HTML, which must include the `g3axonlive.js` script and a call to `G3AxonLive.init()` with the session ID.

---

## Request Lifecycle: Async Update

This happens when the user interacts with a G3AxonLive component (e.g., clicks a button).

```
+----------+      1. User clicks button with      +-----------------+
|          |         data-g3al-id="counter"      |                 |
| Browser  |                                   |                 |
| (JS Side)|      2. g3axonlive.js sends         |                 |
|          |         POST to /g3al/              |                 |
|          |         (includes componentId,      |   AxonASP Server  |
|          |          eventName, sessionId)      | (Go Backend)    |
+----------+ --------------------------------> +-----------------+
     ^                                                  |
     |                                                  | 3. /g3al/ handler:
     |                                                  |    - Looks up session
     |                                                  |    - Re-runs mypage.asp
     |                                                  |
     | 6. g3axonlive.js receives JSON,                  | 4. mypage.asp executes:
     |    swaps element's outerHTML.                    |    - AxonLive.InitPage()
     |    (DOM is updated instantly)                     |      (detects async call)
     |                                                  |    - If AxonLive.IsAsyncRequest...
     |                                                  |    - Handle event
     |                                                  |    - AxonLive.RegisterComponent()
     |                                                  |    - AxonLive.EndAsyncResponse()
     |                                                  |      (sends JSON & halts)
     |                                                  |
     +------------------------------------------------- +
           5. JSON Response
              { "components": [{"id": "counter", "html": "..."}] }
```

1.  A user triggers an event (e.g., `click`) on an HTML element that has `data-g3al-*` attributes.
2.  The `g3axonlive.js` engine intercepts the event, prevents the default browser action, and sends a `fetch` POST request to the `/g3al/` endpoint. The JSON payload contains the session ID, the component ID, the event name, and any arguments.
3.  The Go-based `/g3al/` handler receives the request. It uses the session ID from the payload to look up which ASP page was originally loaded (`mypage.asp`) and re-executes it.
4.  Your ASP script runs again, but this time in an "async context".
    *   `AxonLive.InitPage()` detects this is an async call and parses the event data from the request.
    *   The `AxonLive.IsAsyncRequest` property will now return `True`.
    *   Your code checks for this and executes only the logic needed to handle the specific event (e.g., incrementing a counter).
    *   You call `AxonLive.RegisterComponent()` with the updated HTML for the component.
    *   `AxonLive.EndAsyncResponse()` serializes all registered components into a JSON response, sends it to the client, and immediately halts script execution.
5.  The Go backend sends a lightweight JSON response containing the rendered HTML for the component(s) that changed.
6.  The `g3axonlive.js` engine receives the response, finds the component in the DOM by its ID, and replaces its `outerHTML` with the new content from the server. The update is instantaneous for the user.
