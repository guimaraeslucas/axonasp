# Server-Side API Reference

The `G3AXONLIVE` object provides a clean, procedural API for managing the page lifecycle and handling asynchronous events. You create an instance using `Server.CreateObject("G3AXONLIVE")`.

## Core Lifecycle Methods

These methods form the backbone of every G3AxonLive page.

### InitPage()
Initializes the G3AxonLive lifecycle. This should be the **first method you call** after creating the object. It automatically detects whether the current request is a full page load or an async update.
*   **On a full page load,** it registers the current script's URL against the user's session ID.
*   **On an async update,** it parses the incoming event data (component ID, event name, etc.) and makes it available through other properties.
*   **Returns:** `True` if the request is an async update, `False` otherwise.

```asp
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")
' Returns true on async calls, false on initial load
Dim isAsync
isAsync = AxonLive.InitPage()
```

### RegisterComponent( `componentId`, `html` )
Queues a component's updated HTML to be sent to the client. During an async update, you call this for each component that needs to be redrawn.

*   `componentId` (String): The unique ID of the HTML element to be updated.
*   `html` (String): The full, new `outerHTML` for the component.

```asp
' After updating state, re-render the component and register it
Dim newHtml
newHtml = "<div id=""my-comp"">Updated Content</div>"
AxonLive.RegisterComponent "my-comp", newHtml
```

### GetComponent( `componentId` )
Returns a **Component Proxy Object** that allows you to modify properties, styles, and classes of a specific DOM element directly, without re-rendering its entire HTML.

*   `componentId` (String): The unique ID of the target HTML element.
*   **Returns:** A `G3ALComponentProxy` object.

```asp
Dim btn
Set btn = AxonLive.GetComponent("btnSubmit")
btn.value = "Processing..."
btn.disabled = True
btn.SetStyle "background-color", "gray"
```

### EndAsyncResponse()
Serializes all registered components, granular mutations, and any pending client actions into a single JSON response, sends it to the browser, and immediately halts script execution. This **must** be the final call in your async event handling block.

```asp
If AxonLive.IsAsyncRequest Then
    ' ... handle event, register components or use proxies ...

    ' Send the response and stop
    AxonLive.EndAsyncResponse()
End If
```

## Component Proxy Object (G3ALComponentProxy)

The proxy object returned by `GetComponent` provides a granular API for DOM manipulation.

### Properties
You can read and write standard DOM properties directly on the proxy object. Assignments are queued as client-side actions and also persisted in the `G3ALStore`.

*   `value` (String/Boolean): Sets or gets the `value` property.
*   `disabled` (Boolean): Sets or gets the `disabled` property.
*   `checked` (Boolean): Sets or gets the `checked` property.
*   `className` (String): Sets or gets the `className` property.

### Methods

#### SetStyle( `name`, `value` )
Updates a specific CSS style property on the element.
*   `name` (String): The CSS property name (e.g., "color", "display").
*   `value` (String): The new style value.

#### AddClass( `className` )
Adds a CSS class to the element's `classList`.

#### RemoveClass( `className` )
Removes a CSS class from the element's `classList`.

#### SetAttribute( `name`, `value` )
Sets an HTML attribute on the element.

#### RemoveAttribute( `name` )
Removes an HTML attribute from the element.

#### AddTitle( `title` )
Sets the `title` attribute (tooltip) of the element.

#### RemoveTitle()
Clears the `title` attribute.

#### SetValue( `value` )
Directly sets the `value` property of the element (helper method).

---

## Async Request Properties

These properties are used within your `If AxonLive.IsAsyncRequest` block to get information about the event that triggered the update.

### IsAsyncRequest
*   **(Read-Only Boolean)**
*   Returns `True` if the current script execution was triggered by a G3AxonLive async `fetch` call; otherwise, `False`.

### EventComponentID
*   **(Read-Only String)**
*   The `id` of the component that the user interacted with.

### EventName
*   **(Read-Only String)**
*   The `data-g3al-event-name` of the element that triggered the event.

### GetEventArg( `name` )
*   **Returns:** String
*   Retrieves a single named argument passed from the client. Arguments are defined on the HTML element using `data-g3al-arg-*` attributes.

**Example:**
```html
<button data-g3al-id="my-comp" data-g3al-event="click" data-g3al-event-name="set-value" data-g3al-arg-new-value="42">
    Set to 42
</button>
```
```asp
If AxonLive.EventName = "set-value" Then
    Dim val
    val = AxonLive.GetEventArg("new-value") ' Returns "42"
End If
```

### EventArgs
*   **(Read-Only String)**
*   Returns a JSON-formatted string containing *all* event arguments passed from the client. This is useful for debugging or for parsing with the `G3JSON` library.
