package server

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"net/http"
	"os"
	"strings"
	"time"
)

// MsXML2ServerXMLHTTP implements the MSXML2.ServerXMLHTTP object
// Provides methods for making HTTP requests and handling XML responses
type MsXML2ServerXMLHTTP struct {
	method          string
	url             string
	responseText    string
	responseXML     string
	responseXMLDoc  *MsXML2DOMDocument
	status          int
	statusText      string
	readyState      int
	headers         map[string]string
	responseHeaders map[string]string
	body            string
	responseBody    []byte
	timeout         time.Duration
	async           bool
	ctx             *ExecutionContext
}

// NewMsXML2ServerXMLHTTP creates a new ServerXMLHTTP instance
func NewMsXML2ServerXMLHTTP(ctx *ExecutionContext) *MsXML2ServerXMLHTTP {
	return &MsXML2ServerXMLHTTP{
		headers:         make(map[string]string),
		responseHeaders: make(map[string]string),
		readyState:      0,
		timeout:         30 * time.Second,
		async:           false,
		ctx:             ctx,
	}
}

func (s *MsXML2ServerXMLHTTP) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "responsetext":
		return s.responseText
	case "responsexml":
		if s.responseXMLDoc != nil {
			return s.responseXMLDoc
		}
		return s.responseXML
	case "responsebody":
		if len(s.responseBody) == 0 {
			return NewVBArrayFromValues(0, []interface{}{})
		}
		return bytesToVBArray(s.responseBody)
	case "status":
		return s.status
	case "statustext":
		return s.statusText
	case "readystate":
		return s.readyState
	case "timeout":
		return int(s.timeout.Seconds())
	}
	return nil
}

func (s *MsXML2ServerXMLHTTP) SetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "timeout":
		s.timeout = time.Duration(toInt(value)) * time.Second
	}
	return nil
}

func (s *MsXML2ServerXMLHTTP) GetName() string {
	return "MSXML2.ServerXMLHTTP"
}

func (s *MsXML2ServerXMLHTTP) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch strings.ToLower(name) {
	case "open":
		return s.open(args), nil
	case "setrequestheader":
		return s.setRequestHeader(args), nil
	case "send":
		return s.send(args), nil
	case "abort":
		s.readyState = 4
		return nil, nil
	case "getresponseheader":
		return s.getResponseHeader(args), nil
	case "getallresponseheaders":
		return s.getAllResponseHeaders(), nil
	}
	return nil, nil
}

// open initializes the HTTP request
// Syntax: Open(method, url, [async], [user], [password])
func (s *MsXML2ServerXMLHTTP) open(args []interface{}) interface{} {
	if len(args) < 2 {
		return nil
	}

	s.method = strings.ToUpper(fmt.Sprintf("%v", args[0]))
	s.url = fmt.Sprintf("%v", args[1])

	if len(args) > 2 {
		if async, ok := args[2].(bool); ok {
			s.async = async
		}
	}

	s.readyState = 1
	return nil
}

// setRequestHeader adds a custom header to the request
// Syntax: SetRequestHeader(header, value)
func (s *MsXML2ServerXMLHTTP) setRequestHeader(args []interface{}) interface{} {
	if len(args) < 2 {
		return nil
	}

	key := fmt.Sprintf("%v", args[0])
	value := fmt.Sprintf("%v", args[1])
	s.headers[key] = value
	return nil
}

// send executes the HTTP request
// Syntax: Send([body])
func (s *MsXML2ServerXMLHTTP) send(args []interface{}) interface{} {
	if s.url == "" {
		s.status = 0
		s.statusText = "URL not set"
		s.readyState = 4
		return nil
	}

	s.responseBody = nil
	s.responseXMLDoc = nil
	s.responseText = ""
	s.responseXML = ""

	s.readyState = 2

	var bodyReader io.Reader
	bodyHasContent := false
	bodyIsBinary := false
	if len(args) > 0 && args[0] != nil {
		bodyReader, bodyIsBinary = s.buildRequestBody(args[0])
		bodyHasContent = bodyReader != nil
	}

	req, err := http.NewRequest(s.method, s.url, bodyReader)
	if err != nil {
		s.status = 0
		s.statusText = err.Error()
		s.readyState = 4
		return nil
	}

	// Add custom headers
	for k, v := range s.headers {
		req.Header.Set(k, v)
	}

	// Set default Content-Type if body exists
	if bodyHasContent && req.Header.Get("Content-Type") == "" {
		if bodyIsBinary {
			req.Header.Set("Content-Type", "application/octet-stream")
		} else {
			req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
		}
	}

	s.readyState = 3

	client := &http.Client{Timeout: s.timeout}
	resp, err := client.Do(req)
	if err != nil {
		s.status = 0
		s.statusText = err.Error()
		s.readyState = 4
		return nil
	}
	defer resp.Body.Close()

	// Read response body
	data, err := io.ReadAll(resp.Body)
	if err != nil {
		s.status = resp.StatusCode
		s.statusText = resp.Status
		s.readyState = 4
		return nil
	}

	s.responseBody = data
	s.responseText = string(data)
	s.status = resp.StatusCode
	s.statusText = resp.Status

	// Store response headers
	for k, v := range resp.Header {
		if len(v) > 0 {
			s.responseHeaders[k] = v[0]
		}
	}

	// Parse XML if response is XML
	contentType := resp.Header.Get("Content-Type")
	if s.isXMLResponse(contentType, s.responseText) {
		doc := NewMsXML2DOMDocument(s.ctx)
		if doc != nil {
			doc.loadXML([]interface{}{s.responseText})
			s.responseXMLDoc = doc
		}
		s.responseXML = s.responseText
	}

	s.readyState = 4
	return nil
}

// getResponseHeader retrieves a specific response header
// Syntax: GetResponseHeader(header)
func (s *MsXML2ServerXMLHTTP) getResponseHeader(args []interface{}) interface{} {
	if len(args) < 1 {
		return ""
	}

	key := fmt.Sprintf("%v", args[0])
	if val, ok := s.responseHeaders[key]; ok {
		return val
	}

	// Case-insensitive lookup
	for k, v := range s.responseHeaders {
		if strings.EqualFold(k, key) {
			return v
		}
	}

	return ""
}

// getAllResponseHeaders returns all response headers
func (s *MsXML2ServerXMLHTTP) getAllResponseHeaders() interface{} {
	var result strings.Builder
	for k, v := range s.responseHeaders {
		result.WriteString(fmt.Sprintf("%s: %s\r\n", k, v))
	}
	return result.String()
}

func (s *MsXML2ServerXMLHTTP) buildRequestBody(arg interface{}) (io.Reader, bool) {
	switch v := arg.(type) {
	case *VBArray:
		buf := vbArrayToBytes(v)
		return bytes.NewReader(buf), true
	case []byte:
		return bytes.NewReader(v), true
	default:
		bodyStr := fmt.Sprintf("%v", arg)
		s.body = bodyStr
		return strings.NewReader(bodyStr), false
	}
}

func (s *MsXML2ServerXMLHTTP) isXMLResponse(contentType string, body string) bool {
	if strings.Contains(strings.ToLower(contentType), "xml") {
		return true
	}
	trimmed := strings.TrimSpace(body)
	return strings.HasPrefix(trimmed, "<") && strings.HasSuffix(trimmed, ">")
}

func vbArrayToBytes(arr *VBArray) []byte {
	if arr == nil {
		return nil
	}

	buf := make([]byte, len(arr.Values))
	for i, val := range arr.Values {
		buf[i] = byte(toInt(val))
	}
	return buf
}

func bytesToVBArray(data []byte) *VBArray {
	if len(data) == 0 {
		return NewVBArrayFromValues(0, []interface{}{})
	}

	values := make([]interface{}, len(data))
	for i, b := range data {
		values[i] = int(b)
	}
	return NewVBArrayFromValues(0, values)
}

// ============================================================================
// MsXML2DOMDocument - XML Document Object Model
// ============================================================================

type MsXML2DOMDocument struct {
	xmlContent string
	root       *XMLElement
	async      bool
	parseError *ParseError
	ctx        *ExecutionContext
}

// ParseError represents XML parsing errors
type ParseError struct {
	ErrorCode   int
	ErrorReason string
	FilePos     int
	Line        int
	LinePos     int
	SrcText     string
	URL         string
}

// GetName returns the name of the ParseError object
func (p *ParseError) GetName() string {
	return "IXMLDOMParseError"
}

// GetProperty gets a property from the ParseError
func (p *ParseError) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "errorcode":
		return p.ErrorCode
	case "reason":
		return p.ErrorReason
	case "filepos":
		return p.FilePos
	case "line":
		return p.Line
	case "linepos":
		return p.LinePos
	case "srctext":
		return p.SrcText
	case "url":
		return p.URL
	}
	return nil
}

// SetProperty sets a property on the ParseError (read-only, no-op)
func (p *ParseError) SetProperty(name string, value interface{}) error {
	return nil
}

// CallMethod calls a method on ParseError (none available)
func (p *ParseError) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return nil, nil
}

// XMLElement represents an XML element node
type XMLElement struct {
	Name       string
	Value      string
	Attributes map[string]string
	Children   []*XMLElement
	Parent     *XMLElement
}

// NewMsXML2DOMDocument creates a new DOM Document instance
func NewMsXML2DOMDocument(ctx *ExecutionContext) *MsXML2DOMDocument {
	return &MsXML2DOMDocument{
		async:      false,
		parseError: &ParseError{},
		ctx:        ctx,
	}
}

func (d *MsXML2DOMDocument) GetName() string {
	return "MSXML2.DOMDocument"
}

func (d *MsXML2DOMDocument) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "documentelement":
		// Ensure root is parsed if we have XML content
		if d.root == nil && d.xmlContent != "" {
			if parsed, err := d.parseXMLString(d.xmlContent); err == nil {
				d.root = parsed
			}
		}
		// Return nil (Nothing in VBScript) if no root
		if d.root == nil {
			return nil
		}
		return d.root
	case "xml":
		if d.xmlContent != "" {
			return d.xmlContent
		}
		// If no stored XML but we have a root, generate it
		if d.root != nil {
			return "<?xml version=\"1.0\"?>" + d.elementToXML(d.root, 0)
		}
		return ""
	case "parseerror":
		return d.parseError
	case "async":
		return d.async
	}
	return nil
}

func (d *MsXML2DOMDocument) SetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "async":
		if v, ok := value.(bool); ok {
			d.async = v
		}
	}
	return nil
}

func (d *MsXML2DOMDocument) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch strings.ToLower(name) {
	case "loadxml":
		return d.loadXML(args), nil
	case "load":
		return d.load(args), nil
	case "save":
		return d.save(args), nil
	case "getelementsbytagname":
		return d.getElementsByTagName(args), nil
	case "createelement":
		return d.createElement(args), nil
	case "createtextnode":
		return d.createTextNode(args), nil
	case "createattribute":
		return d.createAttribute(args), nil
	case "appendchild":
		return d.appendChild(args), nil
	case "selectsinglenode":
		return d.selectSingleNode(args), nil
	case "selectnodes":
		return d.selectNodes(args), nil
	}
	return nil, nil
}

// loadXML parses an XML string
// Syntax: LoadXML(xmlString)
func (d *MsXML2DOMDocument) loadXML(args []interface{}) interface{} {
	if len(args) < 1 {
		d.parseError.ErrorCode = -1
		d.parseError.ErrorReason = "No XML provided"
		return false
	}

	xmlStr := fmt.Sprintf("%v", args[0])
	d.xmlContent = xmlStr

	root, err := d.parseXMLString(xmlStr)
	if err != nil || root == nil {
		d.parseError.ErrorCode = -1
		if err != nil {
			d.parseError.ErrorReason = err.Error()
		} else {
			d.parseError.ErrorReason = "Failed to parse XML"
		}
		return false
	}

	d.root = root
	d.parseError.ErrorCode = 0
	d.parseError.ErrorReason = ""
	return true
}

// load loads an XML file from URL or path
// Syntax: Load(url)
func (d *MsXML2DOMDocument) load(args []interface{}) interface{} {
	if len(args) < 1 {
		d.parseError.ErrorCode = -1
		d.parseError.ErrorReason = "No URL provided"
		return false
	}

	urlStr := fmt.Sprintf("%v", args[0])

	// Try to fetch from URL or path
	var content string

	if strings.HasPrefix(urlStr, "http://") || strings.HasPrefix(urlStr, "https://") {
		resp, err := http.Get(urlStr)
		if err != nil {
			d.parseError.ErrorCode = -1
			d.parseError.ErrorReason = err.Error()
			return false
		}
		defer resp.Body.Close()

		data, err := io.ReadAll(resp.Body)
		if err != nil {
			d.parseError.ErrorCode = -1
			d.parseError.ErrorReason = err.Error()
			return false
		}
		content = string(data)
	} else {
		if d.ctx != nil {
			fullPath := d.ctx.Server_MapPath(urlStr)
			data, errFile := getFileContent(fullPath)
			if errFile != nil {
				d.parseError.ErrorCode = -1
				d.parseError.ErrorReason = errFile.Error()
				return false
			}
			content = data
		} else {
			d.parseError.ErrorCode = -1
			d.parseError.ErrorReason = "No context available"
			return false
		}
	}

	d.xmlContent = content
	root, err := d.parseXMLString(content)
	if err != nil || root == nil {
		d.parseError.ErrorCode = -1
		if err != nil {
			d.parseError.ErrorReason = err.Error()
		} else {
			d.parseError.ErrorReason = "Failed to parse XML"
		}
		return false
	}

	d.root = root
	d.parseError.ErrorCode = 0
	d.parseError.ErrorReason = ""
	return true
}

// save saves the XML to a file
// Syntax: Save(filename)
func (d *MsXML2DOMDocument) save(args []interface{}) interface{} {
	if len(args) < 1 {
		return false
	}

	filename := fmt.Sprintf("%v", args[0])
	if d.ctx == nil {
		return false
	}

	fullPath := d.ctx.Server_MapPath(filename)
	content := d.xmlContent
	if d.root != nil {
		content = d.elementToXML(d.root, 0)
	}

	err := saveFileContent(fullPath, content)
	return err == nil
}

// getElementsByTagName finds all elements with a given tag name
// Syntax: GetElementsByTagName(tagName)
func (d *MsXML2DOMDocument) getElementsByTagName(args []interface{}) interface{} {
	if len(args) < 1 {
		// Return empty array instead of nil
		return NewVBArrayFromValues(0, []interface{}{})
	}

	tagName := strings.ToLower(fmt.Sprintf("%v", args[0]))
	var results []*XMLElement

	if d.root != nil {
		d.findElements(d.root, tagName, &results)
	}

	// Convert to interface slice for VBArray
	var interfaceResults []interface{}
	for _, elem := range results {
		interfaceResults = append(interfaceResults, elem)
	}

	// Return as VBArray with 0-based indexing (even if empty)
	return NewVBArrayFromValues(0, interfaceResults)
}

// selectSingleNode finds the first element matching a simple XPath
// Syntax: SelectSingleNode(xpath)
func (d *MsXML2DOMDocument) selectSingleNode(args []interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}

	xpath := fmt.Sprintf("%v", args[0])
	segments, allowAnywhere := tokenizeXPath(xpath)

	if d.root == nil && d.xmlContent != "" {
		if parsed, err := d.parseXMLString(d.xmlContent); err == nil {
			d.root = parsed
		}
	}

	if d.root == nil || len(segments) == 0 {
		return nil
	}

	if allowAnywhere {
		return d.findFirstMatchAnywhere(d.root, segments)
	}

	if !strings.EqualFold(d.root.Name, segments[0]) {
		return nil
	}

	return d.matchFirstFrom(d.root, segments[1:])
}

// selectNodes finds all elements matching a simple XPath
// Syntax: SelectNodes(xpath)
func (d *MsXML2DOMDocument) selectNodes(args []interface{}) interface{} {
	if len(args) < 1 {
		// Return empty array instead of nil
		return NewVBArrayFromValues(0, []interface{}{})
	}

	xpath := fmt.Sprintf("%v", args[0])
	segments, allowAnywhere := tokenizeXPath(xpath)

	if d.root == nil && d.xmlContent != "" {
		if parsed, err := d.parseXMLString(d.xmlContent); err == nil {
			d.root = parsed
		}
	}

	var results []*XMLElement
	if d.root != nil {
		if allowAnywhere {
			d.collectMatchesAnywhere(d.root, segments, &results)
		} else {
			if len(segments) > 0 && strings.EqualFold(d.root.Name, segments[0]) {
				d.collectMatchesFrom(d.root, segments[1:], &results)
			}
		}
	}

	var interfaceResults []interface{}
	for _, elem := range results {
		interfaceResults = append(interfaceResults, elem)
	}

	// Return as VBArray with 0-based indexing (even if empty)
	return NewVBArrayFromValues(0, interfaceResults)
}

// tokenizeXPath returns normalized segments and whether the path should match anywhere (//)
func tokenizeXPath(xpath string) ([]string, bool) {
	trimmed := strings.TrimSpace(xpath)
	allowAnywhere := strings.HasPrefix(trimmed, "//")
	if allowAnywhere {
		trimmed = strings.TrimPrefix(trimmed, "//")
	}

	parts := strings.Split(trimmed, "/")
	var segments []string
	for _, part := range parts {
		seg := strings.TrimSpace(part)
		if seg == "" || seg == "." {
			continue
		}
		segments = append(segments, strings.ToLower(seg))
	}

	return segments, allowAnywhere
}

// matchFirstFrom walks down the tree following the provided path starting at start.
func (d *MsXML2DOMDocument) matchFirstFrom(start *XMLElement, segments []string) *XMLElement {
	if len(segments) == 0 {
		return start
	}

	for _, child := range start.Children {
		if strings.EqualFold(child.Name, segments[0]) {
			if res := d.matchFirstFrom(child, segments[1:]); res != nil {
				return res
			}
		}
	}

	return nil
}

// findFirstMatchAnywhere searches depth-first for the first node that satisfies the path.
func (d *MsXML2DOMDocument) findFirstMatchAnywhere(root *XMLElement, segments []string) *XMLElement {
	if root == nil || len(segments) == 0 {
		return nil
	}

	if strings.EqualFold(root.Name, segments[0]) {
		if res := d.matchFirstFrom(root, segments[1:]); res != nil {
			return res
		}
	}

	for _, child := range root.Children {
		if res := d.findFirstMatchAnywhere(child, segments); res != nil {
			return res
		}
	}

	return nil
}

// collectMatchesFrom gathers all nodes that match the remaining path starting at start.
func (d *MsXML2DOMDocument) collectMatchesFrom(start *XMLElement, segments []string, results *[]*XMLElement) {
	if len(segments) == 0 {
		*results = append(*results, start)
		return
	}

	for _, child := range start.Children {
		if strings.EqualFold(child.Name, segments[0]) {
			d.collectMatchesFrom(child, segments[1:], results)
		}
	}
}

// collectMatchesAnywhere gathers all nodes anywhere in the tree that satisfy the path.
func (d *MsXML2DOMDocument) collectMatchesAnywhere(root *XMLElement, segments []string, results *[]*XMLElement) {
	if root == nil || len(segments) == 0 {
		return
	}

	if strings.EqualFold(root.Name, segments[0]) {
		d.collectMatchesFrom(root, segments[1:], results)
	}

	for _, child := range root.Children {
		d.collectMatchesAnywhere(child, segments, results)
	}
}

// createElement creates a new element
// Syntax: CreateElement(tagName)
func (d *MsXML2DOMDocument) createElement(args []interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}
	tagName := fmt.Sprintf("%v", args[0])

	elem := &XMLElement{
		Name:       tagName,
		Attributes: make(map[string]string),
		Children:   make([]*XMLElement, 0),
	}
	return elem
}

// createTextNode creates a text node
// Syntax: CreateTextNode(text)
func (d *MsXML2DOMDocument) createTextNode(args []interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}

	text := fmt.Sprintf("%v", args[0])
	return &XMLElement{
		Name:  "#text",
		Value: text,
	}
}

// createAttribute creates a new attribute
// Syntax: CreateAttribute(name)
func (d *MsXML2DOMDocument) createAttribute(args []interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}

	name := fmt.Sprintf("%v", args[0])
	attr := &XMLElement{
		Name: name,
	}
	return attr
}

// appendChild adds a child element
// Syntax: AppendChild(newChild)
func (d *MsXML2DOMDocument) appendChild(args []interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}

	if elem, ok := args[0].(*XMLElement); ok {
		if d.root == nil {
			d.root = elem
		} else {
			d.root.Children = append(d.root.Children, elem)
			elem.Parent = d.root
		}
		return elem
	}

	return nil
}

// Helper methods for XMLElement (implements Component interface)
func (e *XMLElement) GetName() string {
	return "IXMLDOMElement"
}

func (e *XMLElement) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "nodename":
		return e.Name
	case "nodevalue":
		return e.Value
	case "text":
		if e.Value != "" {
			return e.Value
		}
		// Concatenate text from all child text nodes
		var text strings.Builder
		for _, child := range e.Children {
			if child.Name == "#text" {
				text.WriteString(child.Value)
			}
		}
		return text.String()
	case "xml":
		// Return XML representation
		return e.toXML(0)
	case "attributes":
		// Return attributes collection
		var attrs []interface{}
		for k, v := range e.Attributes {
			attrs = append(attrs, map[string]interface{}{
				"name":  k,
				"value": v,
			})
		}
		return attrs
	case "childnodes":
		var children []interface{}
		for _, child := range e.Children {
			children = append(children, child)
		}
		return children
	case "firstchild":
		if len(e.Children) > 0 {
			return e.Children[0]
		}
		return nil
	case "lastchild":
		if len(e.Children) > 0 {
			return e.Children[len(e.Children)-1]
		}
		return nil
	case "parentnode":
		return e.Parent
	case "length":
		return len(e.Children)
	case "children":
		// Alias for childnodes
		var children []interface{}
		for _, child := range e.Children {
			children = append(children, child)
		}
		return children
	}
	return nil
}

func (e *XMLElement) SetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "nodevalue":
		e.Value = fmt.Sprintf("%v", value)
	case "text":
		e.Value = fmt.Sprintf("%v", value)
	}
	return nil
}

func (e *XMLElement) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch strings.ToLower(name) {
	case "appendchild":
		if len(args) > 0 {
			if child, ok := args[0].(*XMLElement); ok {
				e.Children = append(e.Children, child)
				child.Parent = e
				return child, nil
			}
		}
	case "getelementsbytagname":
		if len(args) > 0 {
			tagName := strings.ToLower(fmt.Sprintf("%v", args[0]))
			var results []*XMLElement
			e.findElements(tagName, &results)
			var interfaceResults []interface{}
			for _, elem := range results {
				interfaceResults = append(interfaceResults, elem)
			}
			// Return array even if empty
			return NewVBArrayFromValues(0, interfaceResults), nil
		}
	case "item":
		if len(args) > 0 {
			if idx, ok := args[0].(int); ok && idx >= 0 && idx < len(e.Children) {
				return e.Children[idx], nil
			}
		}
	case "setattribute":
		if len(args) >= 2 {
			key := fmt.Sprintf("%v", args[0])
			val := fmt.Sprintf("%v", args[1])
			e.Attributes[key] = val
		}
	case "getattribute":
		if len(args) > 0 {
			key := fmt.Sprintf("%v", args[0])
			return e.Attributes[key], nil
		}
	case "removeattribute":
		if len(args) > 0 {
			key := fmt.Sprintf("%v", args[0])
			delete(e.Attributes, key)
		}
	}
	return nil, nil
}

func (e *XMLElement) findElements(tagName string, results *[]*XMLElement) {
	if strings.ToLower(e.Name) == tagName {
		*results = append(*results, e)
	}
	for _, child := range e.Children {
		child.findElements(tagName, results)
	}
}

func (e *XMLElement) toXML(indent int) string {
	if e.Name == "#text" {
		return e.Value
	}

	var buf bytes.Buffer
	padding := strings.Repeat(" ", indent)

	buf.WriteString(padding + "<" + e.Name)
	for k, v := range e.Attributes {
		buf.WriteString(fmt.Sprintf(` %s="%s"`, k, v))
	}

	if len(e.Children) == 0 && e.Value == "" {
		buf.WriteString(" />\n")
	} else {
		buf.WriteString(">")
		if e.Value != "" {
			buf.WriteString(e.Value)
		}
		if len(e.Children) > 0 {
			buf.WriteString("\n")
			for _, child := range e.Children {
				buf.WriteString(child.toXML(indent + 2))
			}
			buf.WriteString(padding)
		}
		buf.WriteString("</" + e.Name + ">\n")
	}

	return buf.String()
}

// Private helper methods for MsXML2DOMDocument

func (d *MsXML2DOMDocument) parseXMLString(xmlStr string) (*XMLElement, error) {
	trimmed := strings.TrimSpace(xmlStr)
	if trimmed == "" {
		return nil, fmt.Errorf("empty xml")
	}

	decoder := xml.NewDecoder(strings.NewReader(trimmed))
	decoder.Strict = false // be lenient like MSXML

	var root *XMLElement
	stack := make([]*XMLElement, 0)

	for {
		tok, err := decoder.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			node := &XMLElement{
				Name:       t.Name.Local,
				Attributes: make(map[string]string),
				Children:   make([]*XMLElement, 0),
			}
			for _, attr := range t.Attr {
				node.Attributes[attr.Name.Local] = attr.Value
			}

			if len(stack) > 0 {
				parent := stack[len(stack)-1]
				parent.Children = append(parent.Children, node)
				node.Parent = parent
			}

			stack = append(stack, node)
			if root == nil {
				root = node
			}

		case xml.EndElement:
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}

		case xml.CharData:
			text := string(t)
			if strings.TrimSpace(text) == "" {
				continue
			}
			if len(stack) == 0 {
				continue
			}
			parent := stack[len(stack)-1]
			textNode := &XMLElement{
				Name:  "#text",
				Value: text,
			}
			parent.Children = append(parent.Children, textNode)
			textNode.Parent = parent
			if len(parent.Children) == 1 && parent.Value == "" {
				parent.Value = strings.TrimSpace(text)
			}
		}
	}

	return root, nil
}

func (d *MsXML2DOMDocument) findElements(root *XMLElement, tagName string, results *[]*XMLElement) {
	if root == nil {
		return
	}

	if strings.ToLower(root.Name) == tagName {
		*results = append(*results, root)
	}

	for _, child := range root.Children {
		d.findElements(child, tagName, results)
	}
}

func (d *MsXML2DOMDocument) findFirstElement(root *XMLElement, tagName string) *XMLElement {
	if root == nil {
		return nil
	}

	if strings.ToLower(root.Name) == tagName {
		return root
	}

	for _, child := range root.Children {
		if result := d.findFirstElement(child, tagName); result != nil {
			return result
		}
	}

	return nil
}

func (d *MsXML2DOMDocument) elementToXML(elem *XMLElement, indent int) string {
	var buf bytes.Buffer
	padding := strings.Repeat(" ", indent)

	if elem.Name == "#text" {
		return elem.Value
	}

	buf.WriteString(padding + "<" + elem.Name)
	for k, v := range elem.Attributes {
		buf.WriteString(fmt.Sprintf(` %s="%s"`, k, v))
	}

	if len(elem.Children) == 0 && elem.Value == "" {
		buf.WriteString(" />\n")
	} else {
		buf.WriteString(">")
		if elem.Value != "" {
			buf.WriteString(elem.Value)
		}
		if len(elem.Children) > 0 {
			buf.WriteString("\n")
			for _, child := range elem.Children {
				buf.WriteString(d.elementToXML(child, indent+2))
			}
			buf.WriteString(padding)
		}
		buf.WriteString("</" + elem.Name + ">\n")
	}

	return buf.String()
}

// Helper functions for file operations (use OS-level functions)

func getFileContent(path string) (string, error) {
	data, err := os.ReadFile(path)
	return string(data), err
}

func saveFileContent(path string, content string) error {
	return os.WriteFile(path, []byte(content), 0644)
}

// Note: This implementation uses the standard regexp package
// For a more complete regex implementation, you may need to improve parseXMLString
