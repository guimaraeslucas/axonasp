# Client-Side API & Actions

G3AxonLive allows the server to send commands back to the client, instructing it to perform actions beyond simply updating the DOM. These actions are queued using methods on the `G3AXONLIVE` object and are sent to the browser as part of the JSON response from `EndAsyncResponse()`.

The client-side `g3axonlive.js` engine processes these actions automatically.

## Server-Triggered Client Actions

You call these methods from your ASP script during an async update before calling `EndAsyncResponse`.

### Redirect( `url` )
Instructs the browser to navigate to a new page. This is useful for redirecting a user after a form submission or a completed action.

*   `url` (String): The absolute or relative URL to redirect to.

**Example:**
```asp
' In an async update block
If AxonLive.EventName = "save-form" Then
    ' ... process form data ...
    AxonLive.Redirect "/submission-successful.asp"
End If
```

### SetTimer( `componentId`, `eventName`, `delayMs` )
Schedules a future event to be fired on a component. The client will wait for the specified delay and then send a new async request to the server as if the user had triggered the event themselves.

*   `componentId` (String): The ID of the component that will be the target of the future event.
*   `eventName` (String): The name of the event to fire.
*   `delayMs` (Integer): The delay in milliseconds before the event is fired.

**Example: A self-hiding alert message**
```asp
' Show an alert message, then tell the client to fire a "hide" event after 5 seconds
If AxonLive.EventName = "show-alert" Then
    ' Register the visible alert component
    AxonLive.RegisterComponent "alert-box", "<div id='alert-box'>Saved!</div>"
    ' Schedule the 'hide' event
    AxonLive.SetTimer "alert-box", "hide-alert", 5000
End If

' Handle the scheduled event
If AxonLive.EventName = "hide-alert" Then
    ' Register the component as an empty div to hide it
    AxonLive.RegisterComponent "alert-box", "<div id='alert-box'></div>"
End If
```

### Trigger( `componentId`, `eventName` )
Immediately triggers a new async event from the client side without requiring user interaction. This is useful for chaining events.

*   `componentId` (String): The ID of the component to trigger the event on.
*   `eventName` (String): The name of the event to fire.

**Example:**
```asp
' When one action finishes, immediately trigger another
If AxonLive.EventName = "step1" Then
    ' ... do step 1 logic ...
    AxonLive.Trigger "progress-bar", "update-step2"
End If
```

### AddAttribute( `componentId`, `attributeName`, `attributeValue` )
Sets or changes an HTML attribute on a component's element directly in the browser's DOM. This is a lightweight way to make small changes (like modifying a CSS class or disabling a button) without re-rendering the entire component.

*   `componentId` (String): The ID of the element to modify.
*   `attributeName` (String): The name of the attribute (e.g., "class", "disabled").
*   `attributeValue` (String): The new value for the attribute.

**Example: Disabling a button after it's clicked**
```asp
If AxonLive.EventName = "submit-once" Then
    ' Disable the button to prevent double-clicks
    AxonLive.AddAttribute AxonLive.EventComponentID, "disabled", "true"
    ' ... process the submission ...
End If
```

## Manual Client-Side Trigger

You can also trigger events programmatically from your own client-side JavaScript using `G3AxonLive.trigger()`.

### `G3AxonLive.trigger( componentId, eventName, eventArgs )`
Manually fires a G3AxonLive event.

*   `componentId` (String): The target component ID.
*   `eventName` (String): The event name to fire.
*   `eventArgs` (Object): An optional JavaScript object with key-value pairs to be sent as event arguments.

**Example:**
```html
<script>
function startProcess() {
    // Programmatically trigger the 'start' event on the 'process-manager' component
    G3AxonLive.trigger('process-manager', 'start', { mode: 'fast' });
}
</script>
<button onclick="startProcess()">Start a Fast Process</button>
```
This is useful for integrating G3AxonLive with other JavaScript libraries or for creating more complex client-side interactions.
