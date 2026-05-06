# Building Reactive Pages

This guide will walk you through creating your first G3AxonLive page. We will build a simple counter component, demonstrating the core concepts in both VBScript and JScript.

## Step 1: Page Setup

First, create a new `.asp` page. This page needs two essential script blocks in the `<head>`:

1.  **Include `g3axonlive.js`:** This is the client-side engine.
2.  **Initialize the engine:** Call `G3AxonLive.init()` with the current user's `Session.SessionID`.

```html
<!DOCTYPE html>
<html>
<head>
    <title>G3AxonLive Counter</title>
    <script src="/axonlive/g3axonlive.js"></script>
    <script>
        // Initialize the engine with the ASP session ID
        G3AxonLive.init("<%= Session.SessionID %>");
    </script>
</head>
<body>
    <!-- Your components will go here -->
</body>
</html>
```

## Step 2: Create the Server-Side Logic

At the top of your ASP page, you need to set up the G3AxonLive object and handle the page lifecycle.

```asp
<%
' Create the G3AxonLive object
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")

' Initialize the page lifecycle
AxonLive.InitPage()
%>
```

The `InitPage` method is the entry point. It handles both initial page loads and subsequent async updates automatically.

## Step 3: Define a Component

A component is just a piece of HTML with a unique ID and special `data-g3al-*` attributes that link it to the client-side engine.

Let's define our counter component. It will consist of a `div` that displays a number and two buttons to increment and decrement it.

```html
<div id="counter-component">
    <h1>Counter: <span><%= count %></span></h1>
    <button
        data-g3al-id="counter-component"
        data-g3al-event="click"
        data-g3al-event-name="increment">
        + Increment
    </button>
    <button
        data-g3al-id="counter-component"
        data-g3al-event="click"
        data-g3al-event-name="decrement">
        - Decrement
    </button>
</div>
```

**Attribute Breakdown:**

*   `id="counter-component"`: This is the **unique identifier** for our component. It's used by the client engine to know which part of the DOM to update.
*   `data-g3al-id="counter-component"`: This attribute on the buttons tells the engine that they **belong** to the `counter-component`.
*   `data-g3al-event="click"`: Specifies that the engine should intercept the `click` event.
*   `data-g3al-event-name="increment"`: This is the **custom event name** that will be sent to the server. Your server-side code will use this name to decide what to do.

## Step 4: Handle Async Updates

Now, we'll add the server-side logic to handle the `increment` and `decrement` events. The core idea is to check if the current request is an async update, and if so, perform the action and send back the updated component.

---

### VBScript Example

For VBScript, we'll use a `Select Case` block to handle the events.

**`counter_vb.asp`**
```asp
<%@ Language=VBScript %>
<%
' 1. SETUP
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

' 2. STATE MANAGEMENT
Dim count
' On initial load, get count from Session or default to 0
If Not AxonLive.IsAsyncRequest Then
    count = CInt(Session("count"))
    If IsEmpty(count) Or count = "" Then
        count = 0
        Session("count") = count
    End If
' On async update, just get the current value
Else
    count = CInt(Session("count"))
End If

' 3. ASYNC EVENT HANDLING
If AxonLive.IsAsyncRequest Then
    Select Case AxonLive.EventName
        Case "increment"
            count = count + 1
        Case "decrement"
            count = count - 1
    End Select

    ' Save the new state
    Session("count") = count

    ' Re-render the component's HTML into a variable
    Dim updatedHtml
    updatedHtml = "<div id=""counter-component"">" & _
                  "<h1>Counter: <span>" & count & "</span></h1>" & _
                  "<button data-g3al-id=""counter-component"" data-g3al-event=""click"" data-g3al-event-name=""increment"">+ Increment</button> " & _
                  "<button data-g3al-id=""counter-component"" data-g3al-event=""click"" data-g3al-event-name=""decrement"">- Decrement</button>" & _
                  "</div>"

    ' Register the updated HTML for the response
    AxonLive.RegisterComponent "counter-component", updatedHtml

    ' Send the JSON response and stop execution
    AxonLive.EndAsyncResponse()
End If
%>
<!DOCTYPE html>
<html>
<head>
    <title>G3AxonLive Counter (VBScript)</title>
    <script src="/axonlive/g3axonlive.js"></script>
    <script>
        G3AxonLive.init("<%= Session.SessionID %>");
    </script>
</head>
<body>
    <h2>VBScript Counter Example</h2>

    <% ' 4. INITIAL RENDER %>
    <div id="counter-component">
        <h1>Counter: <span><%= count %></span></h1>
        <button data-g3al-id="counter-component" data-g3al-event="click" data-g3al-event-name="increment">+ Increment</button>
        <button data-g3al-id="counter-component" data-g3al-event="click" data-g3al-event-name="decrement">- Decrement</button>
    </div>
</body>
</html>
```

---

### JScript (Server-Side) Example

JScript offers a cleaner, more modern syntax. Using functions can help organize the code better.

**`counter_js.asp`**
```asp
<%@ Language=JScript %>
<%
// 1. SETUP
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

// 2. STATE MANAGEMENT
var count;
if (!AxonLive.IsAsyncRequest) {
    count = parseInt(Session("count"), 10);
    if (isNaN(count)) {
        count = 0;
        Session("count") = count;
    }
} else {
    count = parseInt(Session("count"), 10);
}

// Function to render the component HTML
function renderCounterComponent(currentCount) {
    var html = '<div id="counter-component">';
    html += '<h1>Counter: <span>' + currentCount + '</span></h1>';
    html += '<button data-g3al-id="counter-component" data-g3al-event="click" data-g3al-event-name="increment">+ Increment</button> ';
    html += '<button data-g3al-id="counter-component" data-g3al-event="click" data-g3al-event-name="decrement">- Decrement</button>';
    html += '</div>';
    return html;
}

// 3. ASYNC EVENT HANDLING
if (AxonLive.IsAsyncRequest) {
    switch (String(AxonLive.EventName)) {
        case "increment":
            count++;
            break;
        case "decrement":
            count--;
            break;
    }

    // Save the new state
    Session("count") = count;

    // Register the updated component HTML
    AxonLive.RegisterComponent("counter-component", renderCounterComponent(count));

    // Send JSON response and halt
    AxonLive.EndAsyncResponse();
}
%>
<!DOCTYPE html>
<html>
<head>
    <title>G3AxonLive Counter (JScript)</title>
    <script src="/axonlive/g3axonlive.js"></script>
    <script>
        G3AxonLive.init("<%= Session.SessionID %>");
    </script>
</head>
<body>
    <h2>JScript Counter Example</h2>

    <% // 4. INITIAL RENDER %>
    <%= renderCounterComponent(count) %>
</body>
</html>
```

## VBScript vs. JScript

*   **VBScript** is perfectly viable and will feel familiar to long-time ASP developers. However, string concatenation for building HTML can be cumbersome and error-prone.
*   **JScript** is the recommended approach for new G3AxonLive projects. Its support for functions makes code more modular and reusable (like the `renderCounterComponent` function). It also has better native support for data structures like objects and arrays, which becomes invaluable as your components grow in complexity.
