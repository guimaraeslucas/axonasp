# G3AXON.LIVE Methods

The following methods are available for the `G3AXONLIVE` object.

| Method | Returns | Description |
|---|---|---|
| AddAttribute | Empty | Queues a client action to add or update an HTML attribute on a specific component element. |
| ClearComponentState | Empty | Clears all properties saved in the persistent store for a given session and component pair. |
| EndAsyncResponse | Empty | Serializes all pending HTML patches and client actions into a JSON response, writes it, and halts script execution. |
| GetComponent | Object | Returns a `G3ALComponentProxy` native object for granular DOM manipulation (like `SetStyle` or `AddClass`). |
| GetComponentProperty | String | Retrieves a property value from the persistent global state for a component. |
| GetComponentState | String | Returns a diagnostic string listing all stored properties for a specific component. |
| GetEventArg | String | Retrieves a single named event argument sent by the client. |
| InitPage | Boolean | Parses the incoming request to determine if it is an async G3AxonLive POST. Async context is bound to the authenticated ASP session. Returns `True` if it is. |
| Redirect | Empty | Queues a client action that securely navigates the browser to the specified URL. |
| RegisterComponent | Empty | Queues an HTML patch for a specific component to be included in the async response. |
| RegisterPage | Empty | Records the ASP script URL for a session so the framework knows which page to re-execute for async events. |
| RemoveComponentProperty | Empty | Deletes a specific property entry from the persistent global state. |
| RemoveSession | Empty | Completely deletes all AxonLive state data associated with a specific session ID. |
| SetComponentProperty | Empty | Stores a property value in the persistent global state. |
| SetTimer | Empty | Queues a client action that instructs the browser to trigger a specific event after a set delay in milliseconds. |
| StartCleanup | Empty | Starts a background process to clean up idle AxonLive session data. |
| StopCleanup | Empty | Stops the background cleanup process. |
| Trigger | Empty | Queues a client action that immediately triggers a specific client-side event. |
