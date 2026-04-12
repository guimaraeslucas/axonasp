<%
Option Explicit

Dim t
Dim visiblePasses
Dim visibleFailures

Set t = Server.CreateObject("G3TestSuite")
visiblePasses = 0
visibleFailures = 0

Function HTMLEncode(value)
    Dim text

    If IsObject(value) Then
        If value Is Nothing Then
            text = "Nothing"
        Else
            text = "Object(" & TypeName(value) & ")"
        End If
    ElseIf IsNull(value) Then
        text = "Null"
    ElseIf IsEmpty(value) Then
        text = "Empty"
    Else
        text = CStr(value)
    End If

    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, Chr(34), "&quot;")
    HTMLEncode = text
End Function

Function PreviewValue(value)
    If IsObject(value) Then
        If value Is Nothing Then
            PreviewValue = "Nothing"
        Else
            PreviewValue = "Object(" & TypeName(value) & ")"
        End If
    ElseIf IsNull(value) Then
        PreviewValue = "Null"
    ElseIf IsEmpty(value) Then
        PreviewValue = "Empty"
    Else
        PreviewValue = CStr(value)
    End If
End Function

Function HasUsableObject(value)
    If IsObject(value) Then
        If value Is Nothing Then
            HasUsableObject = False
        Else
            HasUsableObject = True
        End If
    Else
        HasUsableObject = False
    End If
End Function

Sub BeginSection(title)
    Response.Write "<h3>" & HTMLEncode(title) & "</h3>" & vbCrLf
    Response.Write "<ul>" & vbCrLf
    t.Describe title
End Sub

Sub EndSection()
    Response.Write "</ul>" & vbCrLf
End Sub

Sub ReportResult(ok, label, detail)
    If ok Then
        visiblePasses = visiblePasses + 1
        Response.Write "<li><strong>PASS</strong> - " & HTMLEncode(label)
    Else
        visibleFailures = visibleFailures + 1
        Response.Write "<li><strong>FAIL</strong> - " & HTMLEncode(label)
    End If

    If Len(detail) > 0 Then
        Response.Write ": " & HTMLEncode(detail)
    End If

    Response.Write "</li>" & vbCrLf
End Sub

Sub CheckEqual(expected, actual, label)
    Dim ok
    ok = (expected = actual)
    t.AssertEqual expected, actual, label
    Call ReportResult(ok, label, "expected=[" & PreviewValue(expected) & "] actual=[" & PreviewValue(actual) & "]")
End Sub

Sub CheckTrue(condition, label, detail)
    Dim ok
    ok = CBool(condition)
    t.AssertTrue ok, label
    Call ReportResult(ok, label, detail)
End Sub

Sub CheckFalse(condition, label, detail)
    Dim ok
    ok = Not CBool(condition)
    t.AssertFalse condition, label
    Call ReportResult(ok, label, detail)
End Sub

Sub CheckObject(obj, label)
    Dim ok
    ok = HasUsableObject(obj)
    t.AssertTrue ok, label
    Call ReportResult(ok, label, "actual=[" & PreviewValue(obj) & "]")
End Sub

Sub CheckNoError(label)
    Dim currentErr
    Dim currentDesc
    Dim detail

    currentErr = Err.Number
    currentDesc = Err.Description
    t.AssertEqual 0, currentErr, label

    If currentErr = 0 Then
        detail = "Err.Number=0"
    Else
        detail = "Err.Number=" & CStr(currentErr) & " Err.Description=" & currentDesc
    End If

    Call ReportResult(currentErr = 0, label, detail)
    Err.Clear
End Sub

Response.Write "<h1>MSXML2.DOMDocument coverage page</h1>" & vbCrLf
Response.Write "<p>This page exercises the DOMDocument-related surface implemented in lib_msxml.go. Failures identify missing or incompatible behavior.</p>" & vbCrLf

On Error Resume Next

Dim doc
Dim parseError
Dim xmlPayload
Dim root
Dim groupNode
Dim Items
Dim secondItem
Dim names
Dim priceNodes
Dim containsNodes
Dim lastItem
Dim firstItem
Dim unionNodes
Dim nsNodes
Dim attrNodes
Dim textNodes
Dim relativeNodes
Dim followingNodes
Dim precedingNodes
Dim loopNode
Dim loopText
Dim builtDoc
Dim builtRoot
Dim titleNode
Dim titleTextNode
Dim attrNode
Dim reloadedDoc
Dim badDoc
Dim missingDoc
Dim whiteSpaceDoc
Dim whiteSpaceRoot
Dim nextNodeProbe
Dim savePath
Dim loaded
Dim saved
Dim missingLoaded
Dim selectionNamespaces
Dim firstAttr

BeginSection "Object creation and defaults"
Set doc = Server.CreateObject("MSXML2.DOMDocument")
CheckNoError "CreateObject MSXML2.DOMDocument should not raise"
CheckObject doc, "CreateObject should return a DOMDocument object"

Set parseError = doc.parseError
CheckNoError "parseError property access on a new document should not raise"
CheckObject parseError, "parseError should return an object on a new document"

CheckEqual False, doc.async, "Async default should be False"
CheckEqual False, doc.serverHTTPRequest, "ServerHTTPRequest default should be False"
CheckEqual False, doc.resolveExternals, "ResolveExternals default should be False"
CheckEqual False, doc.validateOnParse, "ValidateOnParse default should be False"
CheckEqual False, doc.preserveWhiteSpace, "PreserveWhiteSpace default should be False"
CheckEqual "", doc.selectionLanguage, "SelectionLanguage default should be empty"
CheckEqual "", doc.selectionNamespaces, "SelectionNamespaces default should be empty"
CheckEqual 0, parseError.errorCode, "New document ParseError.ErrorCode should be zero"
CheckEqual "", parseError.reason, "New document ParseError.Reason should be empty"
EndSection

BeginSection "Property mutation and GetProperty/SetProperty"
doc.async = True
CheckNoError "Direct Async property assignment should not raise"
CheckEqual True, doc.async, "Async property should persist after direct assignment"

Call doc.setProperty("ServerHTTPRequest", True)
CheckNoError "SetProperty(ServerHTTPRequest) should not raise"
CheckEqual True, doc.serverHTTPRequest, "ServerHTTPRequest should persist after SetProperty"

doc.resolveExternals = True
CheckNoError "Direct ResolveExternals property assignment should not raise"
CheckEqual True, doc.resolveExternals, "ResolveExternals should persist after direct assignment"

Call doc.setProperty("ValidateOnParse", True)
CheckNoError "SetProperty(ValidateOnParse) should not raise"
CheckEqual True, doc.validateOnParse, "ValidateOnParse should persist after SetProperty"

doc.preserveWhiteSpace = True
CheckNoError "Direct PreserveWhiteSpace assignment should not raise"
CheckEqual True, doc.preserveWhiteSpace, "PreserveWhiteSpace should persist after direct assignment"

Call doc.setProperty("SelectionLanguage", "XPath")
CheckNoError "SetProperty(SelectionLanguage) should not raise"
CheckEqual "XPath", doc.getProperty("SelectionLanguage"), "GetProperty should read back SelectionLanguage"

selectionNamespaces = "xmlns:bk='urn:books'"
doc.selectionNamespaces = selectionNamespaces
CheckNoError "Direct SelectionNamespaces assignment should not raise"
CheckEqual selectionNamespaces, doc.getProperty("SelectionNamespaces"), "GetProperty should read back SelectionNamespaces"
EndSection

BeginSection "LoadXML success path and core DOM traversal"
xmlPayload = "<?xml version=""1.0""?><root><group code=""A""><item id=""1""><name>Alpha</name><price>10</price></item><item id=""2""><name>Beta Guide</name><price>25</price></item><item id=""3""><name>Gamma Guide</name><price>30</price></item></group><bk:book xmlns:bk=""urn:books"" code=""B1"">Namespace Guide</bk:book></root>"
loaded = doc.loadXML(xmlPayload)
CheckNoError "LoadXML(valid) should not raise"
CheckEqual True, loaded, "LoadXML(valid) should return True"

Set root = doc.documentElement
CheckNoError "DocumentElement access after LoadXML should not raise"
CheckObject root, "DocumentElement should return the root node after LoadXML"
CheckEqual "root", root.nodeName, "DocumentElement.NodeName should be root"
CheckTrue InStr(1, doc.xml, "<root>", vbTextCompare) > 0, "XML property should include the root element", Left(doc.xml, 120)

Set groupNode = doc.selectSingleNode("//group")
CheckNoError "SelectSingleNode(//group) should not raise"
CheckObject groupNode, "SelectSingleNode should return the group node"
CheckEqual "A", groupNode.getAttribute("code"), "GetAttribute should read the group code"
CheckEqual 2, root.length, "Root length should expose the number of direct child nodes"
CheckEqual "group", root.firstChild.nodeName, "FirstChild should be the group node"
CheckEqual "book", root.lastChild.nodeName, "LastChild should be the namespaced book node"

Set Items = doc.selectNodes("//item")
CheckNoError "SelectNodes(//item) should not raise"
CheckObject Items, "SelectNodes should return a node list"
CheckEqual 3, Items.length, "NodeList.Length should count all item nodes"
CheckEqual 3, Items.Count, "NodeList.Count should match NodeList.Length"

Set secondItem = Items(1)
CheckNoError "Default NodeList item access should not raise"
CheckObject secondItem, "NodeList default item access should return the second item node"
CheckEqual "item", secondItem.nodeName, "Second item NodeName should be item"
CheckEqual "2", secondItem.getAttribute("id"), "Second item id attribute should be 2"
CheckEqual "root", groupNode.parentNode.nodeName, "ParentNode should point back to the root element"

Set names = doc.getElementsByTagName("name")
CheckNoError "GetElementsByTagName(name) should not raise"
CheckObject names, "GetElementsByTagName should return a node list"
CheckEqual 3, names.length, "GetElementsByTagName(name) should return three nodes"
CheckEqual "Alpha", names(0).text, "The first name node text should be Alpha"

loopText = ""
For Each loopNode In Items
    loopText = loopText & loopNode.selectSingleNode("name").text & "|"
Next
CheckEqual "Alpha|Beta Guide|Gamma Guide|", loopText, "For Each over XMLNodeList should enumerate each node once"

Call secondItem.setAttribute("state", "ready")
CheckNoError "SetAttribute should not raise"
CheckEqual "ready", secondItem.getAttribute("state"), "SetAttribute should persist a new attribute"
Call secondItem.removeAttribute("state")
CheckNoError "RemoveAttribute should not raise"
CheckEqual "", secondItem.getAttribute("state"), "Removed attributes should read back as empty strings"

CheckTrue InStr(1, secondItem.xml, "Beta Guide", vbTextCompare) > 0, "Element.XML should contain element content", secondItem.xml
CheckEqual "Beta Guide", secondItem.selectSingleNode("name").text, "Element.SelectSingleNode should work with relative paths"
EndSection

BeginSection "XPath coverage"
Set priceNodes = doc.selectNodes("//item[price > 20]")
CheckNoError "SelectNodes with numeric child predicate should not raise"
CheckEqual 2, priceNodes.length, "XPath numeric child predicate should filter two items"

Set containsNodes = doc.selectNodes("//item[contains(name, 'Guide')]")
CheckNoError "SelectNodes with contains() predicate should not raise"
CheckEqual 2, containsNodes.length, "contains() predicate should match the guide items"

Set firstItem = doc.selectSingleNode("//item[position() = 1]")
CheckNoError "SelectSingleNode with position() predicate should not raise"
CheckObject firstItem, "position() predicate should return the first item"
CheckEqual "1", firstItem.getAttribute("id"), "position() = 1 should select the first item"

Set lastItem = doc.selectSingleNode("//item[last()]")
CheckNoError "SelectSingleNode with last() predicate should not raise"
CheckObject lastItem, "last() predicate should return the last item"
CheckEqual "3", lastItem.getAttribute("id"), "last() should select the last item"

Set relativeNodes = groupNode.selectNodes("item[starts-with(name, 'Beta')]")
CheckNoError "Relative SelectNodes with starts-with() should not raise"
CheckEqual 1, relativeNodes.length, "starts-with() on a relative XPath should match one item"

Set followingNodes = doc.selectNodes("/root/group/item[1]/following-sibling::item")
CheckNoError "following-sibling axis should not raise"
CheckEqual 2, followingNodes.length, "following-sibling axis should return two items"

Set precedingNodes = doc.selectNodes("/root/group/item[3]/preceding-sibling::item")
CheckNoError "preceding-sibling axis should not raise"
CheckEqual 2, precedingNodes.length, "preceding-sibling axis should return two items"

Set attrNodes = doc.selectNodes("//item/@id")
CheckNoError "Attribute axis selection should not raise"
CheckEqual 3, attrNodes.length, "Attribute axis selection should return one attribute node per item"
Set firstAttr = attrNodes(0)
CheckObject firstAttr, "Attribute axis should produce attribute-like nodes"
CheckEqual "id", firstAttr.nodeName, "Attribute axis node name should be the attribute name"
CheckEqual "1", firstAttr.nodeValue, "Attribute axis node value should be the attribute value"
CheckEqual "item", firstAttr.parentNode.nodeName, "Attribute axis parent should point at the owning item"

Set textNodes = doc.selectNodes("//item/name/text()")
CheckNoError "text() axis selection should not raise"
CheckEqual 3, textNodes.length, "text() selection should return three text nodes"
CheckEqual "#text", textNodes(0).nodeName, "text() selection should return #text nodes"
CheckEqual "Alpha", textNodes(0).nodeValue, "The first text() node should contain Alpha"

Set nsNodes = doc.selectNodes("//bk:book")
CheckNoError "Namespace-aware SelectNodes should not raise"
CheckEqual 1, nsNodes.length, "Namespace-aware XPath should return the namespaced book node"
CheckEqual "Namespace Guide", nsNodes(0).text, "The namespaced book node should expose its text content"

Set unionNodes = doc.selectNodes("//item[1] | //bk:book")
CheckNoError "Union XPath should not raise"
CheckEqual 2, unionNodes.length, "Union XPath should merge distinct results"
EndSection

BeginSection "Document construction and roundtrip load/save"
Set builtDoc = Server.CreateObject("MSXML2.DOMDocument")
CheckNoError "CreateObject for build document should not raise"
CheckObject builtDoc, "Builder document should be created"

Set builtRoot = builtDoc.createElement("generated")
CheckNoError "CreateElement should not raise"
CheckObject builtRoot, "CreateElement should return a node"
Call builtRoot.setAttribute("version", "1")
CheckNoError "SetAttribute on created root should not raise"
CheckEqual "1", builtRoot.getAttribute("version"), "SetAttribute on a created element should persist"

Set titleNode = builtDoc.createElement("title")
Set titleTextNode = builtDoc.createTextNode("Created in ASP")
CheckNoError "CreateTextNode should not raise"
CheckObject titleTextNode, "CreateTextNode should return a text node"
Call titleNode.appendChild(titleTextNode)
CheckNoError "AppendChild should attach a text node to a created element"
CheckEqual "title", titleNode.nodeName, "Created child element NodeName should be title"
CheckEqual "Created in ASP", titleNode.text, "Created child element text should persist"

Call builtRoot.appendChild(titleNode)
CheckNoError "AppendChild should attach the title node to the generated root"
Call builtDoc.appendChild(builtRoot)
CheckNoError "AppendChild on the document should not raise"
CheckEqual "generated", builtDoc.documentElement.nodeName, "The built document should expose the generated root"
CheckEqual "Created in ASP", builtDoc.selectSingleNode("//title").text, "The built document should expose appended title text"

Set attrNode = builtDoc.createAttribute("status")
CheckNoError "CreateAttribute should not raise"
CheckObject attrNode, "CreateAttribute should return an attribute-like node"
attrNode.nodeValue = "draft"
CheckNoError "Assigning NodeValue on a created attribute should not raise"
CheckEqual "status", attrNode.nodeName, "Created attribute NodeName should be status"
CheckEqual "draft", attrNode.nodeValue, "Created attribute NodeValue should persist"

savePath = "runtime_msxml_domdocument_roundtrip.xml"
saved = builtDoc.save(savePath)
CheckNoError "Save should not raise"
CheckEqual True, saved, "Save should return True for a writable relative path"

Set reloadedDoc = Server.CreateObject("MSXML2.DOMDocument")
CheckNoError "CreateObject for reloaded document should not raise"
loaded = reloadedDoc.load(savePath)
CheckNoError "Load(saved file) should not raise"
CheckEqual True, loaded, "Load(saved file) should return True"
If loaded Then
    CheckEqual "generated", reloadedDoc.documentElement.nodeName, "Reloaded document should preserve the root name"
    CheckEqual "Created in ASP", reloadedDoc.selectSingleNode("//title").text, "Reloaded document should preserve appended title text"
Else
    t.AssertTrue False, "Reloaded document checks skipped because Load(saved file) failed"
    Call ReportResult(False, "Reloaded document checks skipped because Load(saved file) failed", "Save/Load roundtrip did not produce a readable file")
End If
EndSection

BeginSection "ParseError handling"
Err.Clear
Set badDoc = Server.CreateObject("MSXML2.DOMDocument")
CheckNoError "CreateObject for bad document should not raise"
loaded = badDoc.loadXML("<root><broken></root>")
CheckNoError "LoadXML(invalid) should not raise a VBScript runtime error"
CheckEqual False, loaded, "LoadXML(invalid) should return False"
Set parseError = badDoc.parseError
CheckObject parseError, "parseError should remain available after invalid LoadXML"
CheckTrue Len(parseError.reason) > 0, "ParseError.Reason should be populated after invalid LoadXML", parseError.reason
CheckTrue parseError.line > 0, "ParseError.Line should be populated after invalid LoadXML", CStr(parseError.line)
CheckTrue parseError.linePos > 0, "ParseError.LinePos should be populated after invalid LoadXML", CStr(parseError.linePos)
CheckTrue InStr(1, parseError.srcText, "<broken>", vbTextCompare) > 0, "ParseError.SrcText should include the invalid source", parseError.srcText
CheckEqual "", parseError.url, "LoadXML parse errors should not set ParseError.URL"

loaded = badDoc.loadXML("<root><ok /></root>")
CheckNoError "LoadXML(valid after failure) should not raise"
CheckEqual True, loaded, "A successful LoadXML after failure should return True"
CheckEqual 0, badDoc.parseError.errorCode, "ParseError.ErrorCode should reset after a successful LoadXML"
CheckEqual "", badDoc.parseError.reason, "ParseError.Reason should reset after a successful LoadXML"
CheckEqual 0, badDoc.parseError.filePos, "ParseError.FilePos should reset after a successful LoadXML"
CheckEqual 0, badDoc.parseError.line, "ParseError.Line should reset after a successful LoadXML"
CheckEqual 0, badDoc.parseError.linePos, "ParseError.LinePos should reset after a successful LoadXML"

Set missingDoc = Server.CreateObject("MSXML2.DOMDocument")
CheckNoError "CreateObject for missing-file document should not raise"
missingLoaded = missingDoc.load("does_not_exist_msxml.xml")
CheckNoError "Load(missing file) should not raise a VBScript runtime error"
CheckEqual False, missingLoaded, "Load(missing file) should return False"
CheckTrue Len(missingDoc.parseError.reason) > 0, "Load(missing file) should populate ParseError.Reason", missingDoc.parseError.reason
CheckEqual "does_not_exist_msxml.xml", missingDoc.parseError.url, "Load(missing file) should populate ParseError.URL with the requested path"
EndSection

BeginSection "Compatibility probes for likely gaps"
Set nextNodeProbe = Items.nextNode()
CheckNoError "NodeList.nextNode probe should not raise"
CheckObject nextNodeProbe, "Compatibility probe: NodeList.nextNode should return the first node before exhaustion"

Set whiteSpaceDoc = Server.CreateObject("MSXML2.DOMDocument")
whiteSpaceDoc.preserveWhiteSpace = True
loaded = whiteSpaceDoc.loadXML("<root>  <a>1</a>  <b>2</b> </root>")
CheckNoError "LoadXML with PreserveWhiteSpace=True should not raise"
CheckEqual True, loaded, "LoadXML with PreserveWhiteSpace=True should still succeed"
Set whiteSpaceRoot = whiteSpaceDoc.documentElement
CheckObject whiteSpaceRoot, "Whitespace probe document should still expose a root node"
CheckEqual "#text", whiteSpaceRoot.firstChild.nodeName, "Compatibility probe: PreserveWhiteSpace=True should retain leading whitespace text nodes"
EndSection

Response.Write "<h2>Visible summary</h2>" & vbCrLf
Response.Write "<p>Checks passed: " & CStr(visiblePasses) & "<br>Checks failed: " & CStr(visibleFailures) & "</p>" & vbCrLf
Response.Write "<p>G3Test totals: total=" & CStr(t.Total) & " passed=" & CStr(t.Passed) & " failed=" & CStr(t.Failed) & "</p>" & vbCrLf

If visibleFailures = 0 And t.Failed = 0 Then
    Response.Write "<p><strong>Result:</strong> no compatibility gaps were exposed by this page.</p>" & vbCrLf
Else
    Response.Write "<p><strong>Result:</strong> one or more checks failed. Review the FAIL lines above to see the current implementation gaps.</p>" & vbCrLf
End If
%>