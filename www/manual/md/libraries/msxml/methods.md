# MSXML2 Methods

## ServerXMLHTTP Methods

| Method | Returns | Description |
|---|---|---|
| `Open(method, url [,async [,user [,password]]])` | Empty | Configures the HTTP method, URL, and optional async flag and credentials. |
| `SetRequestHeader(header, value)` | Empty | Adds or replaces an outgoing request header. |
| `SetTimeouts(resolveTimeout, connectTimeout, sendTimeout, receiveTimeout)` | Empty | Configures DNS resolution, connect, send, and receive timeouts in milliseconds (`0` = infinite). |
| `Send([body])` | Empty | Executes the configured HTTP request. Accepts an optional String or byte-array body. |
| `Abort()` | Empty | Cancels the current request without waiting for a response. |
| `GetResponseHeader(header)` | String | Returns the value of a single response header (case-insensitive lookup). |
| `GetAllResponseHeaders()` | String | Returns all response headers as a CRLF-delimited `Header: Value` string. |

## DOMDocument Methods

| Method | Returns | Description |
|---|---|---|
| `LoadXML(xmlString)` | Boolean | Parses an XML string into the document. Returns True on success. |
| `Load(url)` | Boolean | Loads XML from a file path or HTTP/HTTPS URL. Returns True on success. |
| `Save(filename)` | Boolean | Serializes the document and writes it to a file. Returns True on success. |
| `GetElementsByTagName(tagName)` | XMLNodeList | Returns all descendant elements matching the tag name (case-insensitive). |
| `CreateElement(tagName)` | XMLElement | Creates a new, unattached element node. |
| `CreateTextNode(text)` | XMLElement | Creates a new text node with the supplied string value. |
| `CreateAttribute(name)` | XMLElement | Creates a new, unattached attribute node. |
| `AppendChild(child)` | XMLElement | Attaches a child element to the document root and returns it. |
| `SelectSingleNode(xpath)` | XMLElement | Returns the first node matching the XPath expression, or Null. |
| `SelectNodes(xpath)` | XMLNodeList | Returns all nodes matching the XPath expression. |
| `GetProperty(name)` | Variant | Returns the value of a named document property. |
| `SetProperty(name, value)` | Empty | Sets a named document property (e.g., `SelectionNamespaces`). |

## XMLNodeList Methods

| Method | Returns | Description |
|---|---|---|
| `Item(index)` | XMLElement | Returns the node at the given zero-based index, or Null if out of range. |
| `NextNode()` | XMLElement | Returns the next node in a forward-only iteration, or Null when exhausted. |

## XMLElement Methods

| Method | Returns | Description |
|---|---|---|
| `AppendChild(child)` | XMLElement | Appends a child node to this element and returns the child. |
| `GetElementsByTagName(tagName)` | XMLNodeList | Returns all descendant elements matching the tag name. |
| `Item(index)` | XMLElement | Returns the direct child at the given zero-based index. |
| `SetAttribute(name, value)` | Empty | Creates or replaces an attribute on this element. |
| `GetAttribute(name)` | String | Returns the value of the named attribute, or an empty String if not found. |
| `RemoveAttribute(name)` | Empty | Removes the named attribute from this element. |
| `SelectSingleNode(xpath)` | XMLElement | Returns the first node matching the XPath expression relative to this element, or Null. |
| `SelectNodes(xpath)` | XMLNodeList | Returns all nodes matching the XPath expression relative to this element. |

## Method List by Object

### ServerXMLHTTP
- Open
- SetRequestHeader
- SetTimeouts
- Send
- Abort
- GetResponseHeader
- GetAllResponseHeaders

### DOMDocument
- GetProperty
- SetProperty
- LoadXML
- Load
- Save
- GetElementsByTagName
- CreateElement
- CreateTextNode
- CreateAttribute
- AppendChild
- SelectSingleNode
- SelectNodes

### XMLNodeList
- Item
- NextNode

### XMLElement
- AppendChild
- GetElementsByTagName
- Item
- SetAttribute
- GetAttribute
- RemoveAttribute
- SelectSingleNode
- SelectNodes

## Remarks
- Method names are case-insensitive.
- Runtime behavior follows VM compatibility implementations for each object wrapper.
