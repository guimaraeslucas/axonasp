# Create a Windows COM Library for AxonASP

## Overview
Build a custom AxonASP library that wraps Windows COM objects through go-ole and exposes a clean Classic ASP interface through `Server.CreateObject("<LibraryName>")`.

This guide shows:

- How to split implementation by platform using build tags.
- How to register the object in the VM without reflection.
- How to manage COM memory safely to avoid leaks.
- How to expose `open`, `send`, and `selectNodes` as AxonASP native methods that return `axonvm.Value`.

## Prerequisites

- AxonASP source tree and build environment.
- Go toolchain available on your machine.
- Windows host for COM execution.
- `github.com/go-ole/go-ole` and `github.com/go-ole/go-ole/oleutil` available in the module.
- Basic familiarity with Classic ASP and COM ProgID usage.

## Remarks

<strong>Memory and resource safety is mandatory.</strong> COM wrappers leak quickly when release paths are incomplete. Every object, variant, and COM apartment must be released deterministically. Always use `defer obj.Release()`, `ole.VariantClear(&v)`, and `defer ole.CoUninitialize()` after successful `ole.CoInitialize(0)`.

- COM support is **Windows-only** and must be guarded with build tags.
- Keep dispatch logic switch-based and strongly typed. Do not use reflection.
- Keep dynamic native object routing O(1) by using VM maps keyed by `int64` IDs.
- Return explicit AxonASP values (`NewString`, `NewBool`, `NewInteger`, `Value{Type: VTEmpty}`) and raise AxonASP errors for operational failures.

Although it is tempting to simply use Windows COM objects in AxonASP, understand that memory exhaustion can easily occur. COM objects are legacy technology that Microsoft could deprecate at any moment and will not support forever. Furthermore, using COM objects prevents deployment across other platforms such as Linux, macOS, BSD, and Containers; **this is not the optimal way** to interact with these elements.

**Ideally, you should implement the object as a native GoLang library**. By utilizing AI, converting legacy COM code into high-performance, cross-platform GoLang code is fast and straightforward. Additionally, this allows you to contribute to the AxonASP repository, expanding the server's overall capabilities.

You must perform extensive memory and COM object usage testing before deploying to a production environment. COM implementation poses considerable security risks if implemented incorrectly, as well as the potential to generate denial-of-service errors by overloading the server and causing the application to crash.

Unfortunately, there is no simpler way to implement direct interaction with OLE objects that does not involve creating a native library.

However, since we support `WScript.Shell` and `AxExecute(command)`, nothing prevents you from compiling an external program or running a script in another language and capturing the raw output back into your ASP code. This allows, for example, a program written in any other language to emit content to AxonASP, which can then be read and utilized by the ASP script.

```vbscript
Dim ax, returned
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")
returned = ax.AxExecute(command)
Response.Write "Run returned: " & returned

Dim shell, code
Set shell = Server.CreateObject("WScript.Shell")

code = shell.Run("cmd /c echo G3Pix AxonASP", 0, True)
Response.Write "Run returned: " & code
```
## GoLang Implementation Steps

### Syntax

```asp
Set rss = Server.CreateObject("MSXML.COM")
Call rss.open("GET", "http://rss.slashdot.org/Slashdot/slashdot", False)
Call rss.send()
result = rss.selectNodes("/rdf:RDF/item")
```

### Parameters and Arguments

| Method | Parameter | Type | Required | Description |
|---|---|---|---|---|
| `open` | `method` | String | Yes | HTTP verb. Use `GET` for RSS fetch scenarios. |
| `open` | `url` | String | Yes | Absolute URL requested through `Microsoft.XMLHTTP`. |
| `open` | `async` | Boolean | No | Async mode. Use `False` for deterministic Classic ASP flow. |
| `send` | `body` | String/Empty | No | Optional request body. Use `Empty` for GET requests. |
| `selectNodes` | `xpath` | String | Yes | XPath expression evaluated against `responseXml`. |

### Return Values

- `open`: Returns `Boolean`.
  `True` when COM call succeeds and request metadata is configured; `False` when validation or COM call fails.
- `send`: Returns `Boolean`.
  `True` when the request reaches `readyState = 4`; `False` on timeout or COM failure.
- `selectNodes`: Returns `String`.
  CRLF-delimited payload where each line is `title|link` for one matched node.
- `LastError` (property get): Returns `String` with the latest failure message, or an empty string when no error exists.

### 1. Add the Windows implementation file

Create `axonvm/lib_msxmlcom_windows.go` with this build tag:

```go
//go:build windows && !lib_msxmlcom_disabled
```

Example implementation skeleton:

```go
//go:build windows && !lib_msxmlcom_disabled

package axonvm

import (
    "fmt"
    "runtime"
    "strings"
    "time"

    "github.com/go-ole/go-ole"
    "github.com/go-ole/go-ole/oleutil"
)

// MSXMLCOM exposes Microsoft.XMLHTTP + responseXml.selectNodes through AxonASP.
type MSXMLCOM struct {
    vm         *VM
    xhr        *ole.IDispatch
    responseXM *ole.IDispatch
    lastError  string
}

// newMSXMLCOMObject allocates and registers one native object ID.
func (vm *VM) newMSXMLCOMObject() Value {
    obj := &MSXMLCOM{vm: vm}
    id := vm.nextDynamicNativeID
    vm.nextDynamicNativeID++
    vm.msxmlCOMItems[id] = obj
    return Value{Type: VTNativeObject, Num: id}
}

func (m *MSXMLCOM) setError(prefix string, err error) Value {
    if err != nil {
        m.lastError = fmt.Sprintf("%s: %v", prefix, err)
    } else {
        m.lastError = prefix
    }
    return NewBool(false)
}

func (m *MSXMLCOM) DispatchPropertyGet(propertyName string) Value {
    switch {
    case strings.EqualFold(propertyName, "LastError"):
        return NewString(m.lastError)
    }
    return Value{Type: VTEmpty}
}

func (m *MSXMLCOM) DispatchPropertySet(propertyName string, args []Value) bool {
    return false
}

func (m *MSXMLCOM) DispatchMethod(methodName string, args []Value) Value {
    switch {
    case strings.EqualFold(methodName, "open"):
        if len(args) < 2 {
            return m.setError("MSXML.COM.open requires method and url", nil)
        }
        method := args[0].String()
        url := args[1].String()
        async := false
        if len(args) >= 3 {
            async = m.vm.asBool(args[2])
        }

        if m.xhr == nil {
            runtime.LockOSThread()
            if err := ole.CoInitialize(0); err != nil {
                runtime.UnlockOSThread()
                return m.setError("CoInitialize failed", err)
            }

            unknown, err := oleutil.CreateObject("Microsoft.XMLHTTP")
            if err != nil {
                ole.CoUninitialize()
                runtime.UnlockOSThread()
                return m.setError("CreateObject(Microsoft.XMLHTTP) failed", err)
            }
            defer unknown.Release()

            xhr, err := unknown.QueryInterface(ole.IID_IDispatch)
            if err != nil {
                ole.CoUninitialize()
                runtime.UnlockOSThread()
                return m.setError("QueryInterface(IDispatch) failed", err)
            }
            m.xhr = xhr
        }

        v, err := oleutil.CallMethod(m.xhr, "open", method, url, async)
        if err != nil {
            return m.setError("XMLHTTP.open failed", err)
        }
        ole.VariantClear(&v)
        m.lastError = ""
        return NewBool(true)

    case strings.EqualFold(methodName, "send"):
        if m.xhr == nil {
            return m.setError("XMLHTTP is not initialized. Call open first", nil)
        }

        var v ole.VARIANT
        var err error
        if len(args) >= 1 && args[0].Type != VTEmpty {
            v, err = oleutil.CallMethod(m.xhr, "send", args[0].String())
        } else {
            v, err = oleutil.CallMethod(m.xhr, "send")
        }
        if err != nil {
            return m.setError("XMLHTTP.send failed", err)
        }
        ole.VariantClear(&v)

        deadline := time.Now().Add(30 * time.Second)
        for {
            stateVar, stateErr := oleutil.GetProperty(m.xhr, "readyState")
            if stateErr != nil {
                return m.setError("XMLHTTP.readyState failed", stateErr)
            }
            state := int32(stateVar.Val)
            ole.VariantClear(&stateVar)
            if state == 4 {
                break
            }
            if time.Now().After(deadline) {
                return m.setError("XMLHTTP.send timed out waiting for readyState=4", nil)
            }
            time.Sleep(5 * time.Millisecond)
        }

        if m.responseXM != nil {
            m.responseXM.Release()
            m.responseXM = nil
        }

        xmlVar, xmlErr := oleutil.GetProperty(m.xhr, "responseXml")
        if xmlErr != nil {
            return m.setError("XMLHTTP.responseXml failed", xmlErr)
        }

        respDisp := xmlVar.ToIDispatch()
        ole.VariantClear(&xmlVar)
        if respDisp == nil {
            return m.setError("responseXml returned nil", nil)
        }
        m.responseXM = respDisp
        m.lastError = ""
        return NewBool(true)

    case strings.EqualFold(methodName, "selectNodes"):
        if m.responseXM == nil {
            return m.setError("responseXml is not available. Call send first", nil)
        }
        if len(args) < 1 {
            return m.setError("selectNodes requires xpath", nil)
        }

        nodeVar, err := oleutil.CallMethod(m.responseXM, "selectNodes", args[0].String())
        if err != nil {
            return m.setError("responseXml.selectNodes failed", err)
        }
        nodes := nodeVar.ToIDispatch()
        ole.VariantClear(&nodeVar)
        if nodes == nil {
            return NewString("")
        }
        defer nodes.Release()

        countVar, countErr := oleutil.GetProperty(nodes, "length")
        if countErr != nil {
            return m.setError("NodeList.length failed", countErr)
        }
        count := int(countVar.Val)
        ole.VariantClear(&countVar)

        var sb strings.Builder
        for i := 0; i < count; i++ {
            itemVar, itemErr := oleutil.CallMethod(nodes, "item", i)
            if itemErr != nil {
                return m.setError("NodeList.item failed", itemErr)
            }
            item := itemVar.ToIDispatch()
            ole.VariantClear(&itemVar)
            if item == nil {
                continue
            }

            titleVar, _ := oleutil.CallMethod(item, "selectSingleNode", "title")
            titleNode := titleVar.ToIDispatch()
            ole.VariantClear(&titleVar)

            linkVar, _ := oleutil.CallMethod(item, "selectSingleNode", "link")
            linkNode := linkVar.ToIDispatch()
            ole.VariantClear(&linkVar)

            title := ""
            if titleNode != nil {
                txtVar, _ := oleutil.GetProperty(titleNode, "text")
                title = txtVar.ToString()
                ole.VariantClear(&txtVar)
                titleNode.Release()
            }

            link := ""
            if linkNode != nil {
                txtVar, _ := oleutil.GetProperty(linkNode, "text")
                link = txtVar.ToString()
                ole.VariantClear(&txtVar)
                linkNode.Release()
            }

            item.Release()

            if i > 0 {
                sb.WriteString("\r\n")
            }
            sb.WriteString(title)
            sb.WriteString("|")
            sb.WriteString(link)
        }

        m.lastError = ""
        return NewString(sb.String())

    case strings.EqualFold(methodName, "close"):
        if m.responseXM != nil {
            m.responseXM.Release()
            m.responseXM = nil
        }
        if m.xhr != nil {
            m.xhr.Release()
            m.xhr = nil
            ole.CoUninitialize()
            runtime.UnlockOSThread()
        }
        m.lastError = ""
        return Value{Type: VTEmpty}
    }

    return Value{Type: VTEmpty}
}
```

One-shot COM calls should still use full defer-based teardown:

```go
func fetchStatusCodeOnce(url string) (int32, error) {
    runtime.LockOSThread()
    defer runtime.UnlockOSThread()

    if err := ole.CoInitialize(0); err != nil {
        return 0, err
    }
    defer ole.CoUninitialize()

    unknown, err := oleutil.CreateObject("Microsoft.XMLHTTP")
    if err != nil {
        return 0, err
    }
    defer unknown.Release()

    xhr, err := unknown.QueryInterface(ole.IID_IDispatch)
    if err != nil {
        return 0, err
    }
    defer xhr.Release()

    openVar, err := oleutil.CallMethod(xhr, "open", "GET", url, false)
    if err != nil {
        return 0, err
    }
    ole.VariantClear(&openVar)

    sendVar, err := oleutil.CallMethod(xhr, "send")
    if err != nil {
        return 0, err
    }
    ole.VariantClear(&sendVar)

    statusVar, err := oleutil.GetProperty(xhr, "status")
    if err != nil {
        return 0, err
    }
    defer ole.VariantClear(&statusVar)

    return int32(statusVar.Val), nil
}
```

### 2. Add the non-Windows or disabled fallback file

Create `axonvm/lib_msxmlcom_disabled.go` with this build tag:

```go
//go:build !windows || lib_msxmlcom_disabled
```

Use a disabled stub that fails gracefully:

```go
//go:build !windows || lib_msxmlcom_disabled

package axonvm

type MSXMLCOM struct{}

func (vm *VM) newMSXMLCOMObject() Value {
    panicLibraryDisabled("msxmlcom", "MSXML.COM library")
    return Value{Type: VTEmpty}
}

func (m *MSXMLCOM) DispatchPropertyGet(propertyName string) Value              { return Value{Type: VTEmpty} }
func (m *MSXMLCOM) DispatchPropertySet(propertyName string, args []Value) bool { return false }
func (m *MSXMLCOM) DispatchMethod(methodName string, args []Value) Value       { return Value{Type: VTEmpty} }
```

This pattern avoids compilation/runtime breakage on Linux, macOS, and WASM while returning a standard AxonASP disabled-library error.

### 3. Register the object in the VM without reflection

Update `axonvm/vm.go` in the same style used by existing native libraries.

1. Add one object map to the VM struct:

```go
msxmlCOMItems map[int64]*MSXMLCOM
```

2. Initialize the map in `NewVM`:

```go
msxmlCOMItems: make(map[int64]*MSXMLCOM),
```

3. Register the ProgID in `Server.CreateObject` dispatch:

```go
if progIDKey == "msxml.com" {
    return vm.newMSXMLCOMObject()
}
```

4. Route method calls in `dispatchNativeCall`:

```go
if msxmlCOMObject, exists := vm.msxmlCOMItems[objID]; exists {
    return msxmlCOMObject.DispatchMethod(member, args)
}
```

5. Route property gets in member dispatch:

```go
if msxmlCOMObject, exists := vm.msxmlCOMItems[target.Num]; exists {
    return msxmlCOMObject.DispatchPropertyGet(member)
}
```

6. Route property sets where native setters are handled:

```go
if msxmlCOMObject, exists := vm.msxmlCOMItems[target.Num]; exists {
    return msxmlCOMObject.DispatchPropertySet(member, args)
}
```

No reflection is required. All routing remains explicit, deterministic, and allocation-aware.

### 4. Add cleanup hooks for request teardown

If the library instance can outlive one method call, add deterministic cleanup paths where VM native resources are reset.

```go
for id, obj := range vm.msxmlCOMItems {
    _ = obj.DispatchMethod("close", nil)
    delete(vm.msxmlCOMItems, id)
}
```

This prevents COM handles from surviving request boundaries.

## Code Example

Classic ASP usage with `Server.CreateObject("MSXML.COM")`:

```asp
<%
Option Explicit

Dim rss, raw, lines, i, pair, titleText, linkText

Set rss = Server.CreateObject("MSXML.COM")

If Not rss.open("GET", "http://rss.slashdot.org/Slashdot/slashdot", False) Then
    Response.Write "Open failed: " & rss.LastError
    Set rss = Nothing
    Response.End
End If

If Not rss.send() Then
    Response.Write "Send failed: " & rss.LastError
    Set rss = Nothing
    Response.End
End If

raw = rss.selectNodes("/rdf:RDF/item")
lines = Split(raw, vbCrLf)

Response.Write "<h2>RSS Items</h2>"
Response.Write "<ul>"

For i = 0 To UBound(lines)
    If Len(lines(i)) > 0 Then
        pair = Split(lines(i), "|")
        titleText = ""
        linkText = ""

        If UBound(pair) >= 0 Then titleText = pair(0)
        If UBound(pair) >= 1 Then linkText = pair(1)

        Response.Write "<li><a href=\"" & Server.HTMLEncode(linkText) & "\">" & Server.HTMLEncode(titleText) & "</a></li>"
    End If
Next

Response.Write "</ul>"

Call rss.close()
Set rss = Nothing
%>
```

## API Reference

### Native object contract

| Member | Access | AxonASP type | Description |
|---|---|---|---|
| `open(method, url, async)` | Method | `Boolean` | Initializes COM apartment if needed and configures request metadata. |
| `send([body])` | Method | `Boolean` | Sends request and waits for `readyState = 4` with timeout guard. |
| `selectNodes(xpath)` | Method | `String` | Extracts XML nodes and returns `title|link` entries separated by CRLF. |
| `close()` | Method | `Empty` | Releases COM dispatch objects and uninitializes COM apartment. |
| `LastError` | Property Get | `String` | Most recent operational error generated by the wrapper. |

### Required release operations

- `defer obj.Release()` for every COM `IDispatch` or `IUnknown` acquisition.
- `ole.VariantClear(&v)` for every `ole.VARIANT` returned by `GetProperty` or `CallMethod`.
- `defer ole.CoUninitialize()` or equivalent deterministic close-path uninitialization.
- `runtime.LockOSThread()` when COM apartment lifetime must stay on one OS thread.

### Build tags

- Windows enabled file: `//go:build windows && !lib_msxmlcom_disabled`
- Disabled fallback file: `//go:build !windows || lib_msxmlcom_disabled`
