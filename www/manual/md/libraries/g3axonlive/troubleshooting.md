# Troubleshooting

Here are solutions to some common issues you might encounter while developing with G3AxonLive.

---
### **Q: Why isn't my event firing? My button click does nothing.**

**A:** This is usually caused by one of a few common setup errors. Check them in this order:

1.  **Is `g3axonlive.js` included?**
    Make sure your page's `<head>` includes `<script src="/axonlive/g3axonlive.js"></script>`.

2.  **Is `G3AxonLive.init()` being called?**
    Immediately after including the script, you MUST initialize it with the session ID:
    `<script>G3AxonLive.init("<%= Session.SessionID %>");</script>`
    If the session ID is missing or incorrect, the client engine cannot communicate with the server.

3.  **Do the HTML attributes match?**
    *   The interactive element (e.g., `<button>`) must have a `data-g3al-id` attribute that points to the **ID of the component's container element**.
    *   The container `<div>` must have a matching `id`.

    *Correct:*
    ```html
    <div id="my-component">
        <button data-g3al-id="my-component" ... >Click Me</button>
    </div>
    ```

    *Incorrect:*
    ```html
    <div id="my-component">
         <!-- This button's g3al-id points to itself, not the container -->
        <button id="my-button" data-g3al-id="my-button" ... >Click Me</button>
    </div>
    ```

4.  **Is there a JavaScript error?**
    Open your browser's developer console (F12). The `g3axonlive.js` engine logs errors and warnings there. Look for messages like "G3AxonLive: sessionId is required" or "G3AxonLive: Component element not found".

---
### **Q: How do I debug the JSON patch response?**

**A:** The communication between the client and server happens via `fetch` requests. You can inspect these using your browser's developer tools.

1.  Open the developer tools (F12) and go to the **Network** tab.
2.  Filter the requests by "Fetch/XHR".
3.  Trigger an event in your application. You will see a new request appear, likely to the `/g3al/` endpoint.
4.  Click on this request.
    *   The **Headers** tab shows the request details, including the `X-G3AxonLive: true` header.
    *   The **Payload** or **Request** tab shows the JSON data sent *to* the server (e.g., `{ "componentId": "...", "eventName": "..." }`).
    *   The **Response** tab shows the JSON data sent *from* the server. This is the most important part. You should see a structure like this:
        ```json
        {
            "success": true,
            "components": [
                {
                    "componentId": "counter-component",
                    "html": "<div id="counter-component">...</div>"
                }
            ],
            "actions": []
        }
        ```
    If `success` is `false`, the `error` property will contain a server-side error message. If the `html` looks wrong, there's a bug in your server-side rendering logic.

---
### **Q: My component updates, but then the events stop working.**

**A:** This happens when the new HTML you register with `AxonLive.RegisterComponent` is missing the required `data-g3al-*` attributes on its interactive elements.

When G3AxonLive replaces the component's `outerHTML`, the old elements and their event listeners are destroyed. The new elements must have the same `data-g3al-*` attributes so the client engine can re-bind the events.

**Always ensure your server-side rendering logic (whether in VBScript or a JScript function) generates the complete and correct HTML for the component in all its states.**

---
### **Q: What do the error codes mean?**

**A:** G3AxonLive uses specific error codes for common failures. These are raised as standard AxonASP errors.

*   **`ErrG3ALComponentLimitExceeded` (Code 2100):** You tried to call `RegisterComponent` too many times in a single async request. The default limit is 200 components per response. This is a safeguard against performance issues.
*   **`ErrG3ALInvalidComponentID` (Code 2101):** The `componentId` you provided contains invalid characters. IDs should only contain letters, numbers, hyphens, underscores, dots, and colons.
*   **`ErrG3ALResponseAlreadyEnded` (Code 2102):** You called `EndAsyncResponse` more than once in the same request. This method halts execution, so any code after it will not run.
*   **`ErrG3ALTimerDelayInvalid` (Code 2103):** The `delayMs` passed to `SetTimer` was less than or equal to zero. The delay must be a positive integer.
