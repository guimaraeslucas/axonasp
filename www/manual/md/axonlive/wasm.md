# G3AxonLive in WebAssembly (WASM)

## Overview
AxonASP WASM supports **AxonLive**, the native reactive component framework. Unlike the server-side implementation which uses HTTP JSON payloads and client-side JavaScript polling to apply patches, the WASM implementation acts much like **Microsoft Blazor**.

When the AxonASP engine runs entirely inside the browser's WebAssembly sandbox, `G3AXONLIVE` bypasses all network boundaries. Method calls in Classic ASP (like `SetStyle`, `AddClass`, or `SetValue`) are directly translated into real-time DOM manipulations via the `syscall/js` API.

This enables you to write highly interactive, Single-Page Applications (SPAs) entirely in VBScript or Server-Side JavaScript.

## Compatibility
The API of `G3AXONLIVE` in WASM is **100% compatible** with the server-side version. Code written for the server will run seamlessly in the browser.

However, in the WASM environment:
1. **Zero Network Latency**: UI changes happen instantly in the same browser tick.
2. **In-Memory State**: Component states and session variables are stored in the browser's memory, ensuring blazing-fast reads and writes.
3. **True Event Driven**: Methods like `SetTimer` use native browser `setTimeout`, making the ASP engine yield to the browser's event loop asynchronously.

## WASM Integration Architecture
To handle AxonLive events from the browser DOM into the WASM runtime, AxonASP exports the `AxonASP.dispatchLiveEvent` function.

### Setting up the JS Shim
In your HTML page that hosts the WASM module, you need to map DOM events to the WASM engine.

```html
<script>
    // Global shim for AxonLive in WASM
    window.G3AxonLive = {
        currentCode: "", // The raw ASP code currently executing
        dispatch: async function(componentId, eventName, eventArgs = {}) {
            if (typeof AxonASP !== "undefined") {
                try {
                    // Send the event directly to the WASM runtime
                    await AxonASP.dispatchLiveEvent("wasm-session", componentId, eventName, JSON.stringify(eventArgs), this.currentCode);
                } catch (e) {
                    console.error("AxonLive WASM Error:", e);
                }
            }
        }
    };

    // Helper to map HTML onclick="" attributes
    window.dispatchLiveEvent = function(componentId, eventName, eventArgs = {}) {
        window.G3AxonLive.dispatch(componentId, eventName, eventArgs);
    };
</script>
```

### Example: A Reactive Counter in WASM

This is a complete example of a reactive component running entirely in the browser using VBScript and WebAssembly.

```vbscript
<%
Dim Axon
Set Axon = Server.CreateObject("G3AXONLIVE")

If Axon.InitPage() Then
    ' We are handling an async WASM event!
    If Axon.EventName = "click" Then
        Dim currentCount
        currentCount = Axon.GetComponentProperty("btnCounter", "data-count")
        
        If IsEmpty(currentCount) Then currentCount = 0
        currentCount = CInt(currentCount) + 1
        
        ' Update internal state
        Axon.SetComponentProperty "btnCounter", "data-count", currentCount
        
        ' Directly manipulate the DOM via WASM syscall/js!
        Axon.GetComponent("btnCounter").SetStyle "background", "#3366cc"
        Axon.GetComponent("btnCounter").SetStyle "color", "white"
        Axon.GetComponent("btnCounter").SetValue "Clicked " & currentCount & " times!"
    End If
    
    Axon.EndAsyncResponse()
End If

' Initial Page Render
Response.Write "<h1>WASM Reactive Counter</h1>"
Response.Write "<button id='btnCounter' onclick='dispatchLiveEvent(""btnCounter"", ""click"")'>Click Me</button>"
%>
```

## Supported Methods
All component manipulation methods are fully supported in WASM and immediately applied to the browser's DOM element matching the `ComponentID`:

- `SetStyle(property, value)`
- `AddClass(className)`
- `RemoveClass(className)`
- `SetAttribute(attribute, value)`
- `RemoveAttribute(attribute)`
- `SetValue(value)` -> Updates `innerHTML` or input `value`.
- `AddTitle(title)`
- `RemoveTitle()`

## Considerations
- **Source Code Resolution**: Unlike a server, WASM does not have a physical file system. To re-compile the ASP script upon an event, the JS shim must pass the original ASP source code to `dispatchLiveEvent`.
- **Async Execution**: The `execute` and `dispatchLiveEvent` JS exports return **Promises**. You must `await` them if you are managing UI loading states.
