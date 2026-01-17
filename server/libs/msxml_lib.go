package asp

import (
	"bytes"
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
	status          int
	statusText      string
	readyState      int
	headers         map[string]string
	responseHeaders map[string]string
	body            string
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
		return s.responseXML
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

func (s *MsXML2ServerXMLHTTP) SetProperty(name string, value interface{}) {
	switch strings.ToLower(name) {
	case "timeout":
		if v, ok := value.(int); ok {
			s.timeout = time.Duration(v) * time.Second
		}
	}
}

func (s *MsXML2ServerXMLHTTP) CallMethod(name string, args []interface{}) interface{} {
	switch strings.ToLower(name) {
	case "open":
		return s.open(args)
	case "setrequestheader":
		return s.setRequestHeader(args)
	case "send":
		return s.send(args)
	case "abort":
		s.readyState = 4
		return nil
	case "getresponseheader":
		return s.getResponseHeader(args)
	case "getallresponseheaders":
		return s.getAllResponseHeaders()
	}
	return nil
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

	s.readyState = 2

	var bodyReader io.Reader
	if len(args) > 0 && args[0] != nil {
		bodyStr := fmt.Sprintf("%v", args[0])
		bodyReader = strings.NewReader(bodyStr)
		s.body = bodyStr
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
	if s.body != "" && req.Header.Get("Content-Type") == "" {
		req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
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
	if strings.Contains(strings.ToLower(contentType), "xml") ||
		strings.HasPrefix(strings.TrimSpace(s.responseText), "<") {
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

func (d *MsXML2DOMDocument) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "documentelement":
		return d.root
	case "xml":
		return d.xmlContent
	case "parseerror":
		return d.parseError
	case "async":
		return d.async
	}
	return nil
}

func (d *MsXML2DOMDocument) SetProperty(name string, value interface{}) {
	switch strings.ToLower(name) {
	case "async":
		if v, ok := value.(bool); ok {
			d.async = v
		}
	}
}

func (d *MsXML2DOMDocument) CallMethod(name string, args []interface{}) interface{} {
	switch strings.ToLower(name) {
	case "loadxml":
		return d.loadXML(args)
	case "load":
		return d.load(args)
	case "save":
		return d.save(args)
	case "getelementsbytagname":
		return d.getElementsByTagName(args)
	case "createelement":
		return d.createElement(args)
	case "createtextnode":
		return d.createTextNode(args)
	case "createattribute":
		return d.createAttribute(args)
	case "appendchild":
		return d.appendChild(args)
	case "selectsinglenode":
		return d.selectSingleNode(args)
	case "selectnodes":
		return d.selectNodes(args)
	}
	return nil
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

	// Simple XML parsing - build element tree
	d.root = d.parseXMLString(xmlStr)
	if d.root == nil {
		d.parseError.ErrorCode = -1
		d.parseError.ErrorReason = "Failed to parse XML"
		return false
	}

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

	// Try to fetch from URL or file
	var content string

	if strings.HasPrefix(urlStr, "http://") || strings.HasPrefix(urlStr, "https://") {
		// HTTP request
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
		// Local file
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
	d.root = d.parseXMLString(content)

	if d.root == nil {
		d.parseError.ErrorCode = -1
		d.parseError.ErrorReason = "Failed to parse XML"
		return false
	}

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
		return []interface{}{}
	}

	tagName := strings.ToLower(fmt.Sprintf("%v", args[0]))
	var results []*XMLElement

	if d.root != nil {
		d.findElements(d.root, tagName, &results)
	}

	// Convert to interface slice
	var interfaceResults []interface{}
	for _, elem := range results {
		interfaceResults = append(interfaceResults, elem)
	}

	return interfaceResults
}

// selectSingleNode finds the first element matching a simple XPath
// Syntax: SelectSingleNode(xpath)
func (d *MsXML2DOMDocument) selectSingleNode(args []interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}

	xpath := fmt.Sprintf("%v", args[0])
	// Simple XPath support: //tagname or /root/child
	tagName := d.extractTagFromXPath(xpath)

	if d.root != nil {
		elem := d.findFirstElement(d.root, tagName)
		return elem
	}

	return nil
}

// selectNodes finds all elements matching a simple XPath
// Syntax: SelectNodes(xpath)
func (d *MsXML2DOMDocument) selectNodes(args []interface{}) interface{} {
	if len(args) < 1 {
		return []interface{}{}
	}

	xpath := fmt.Sprintf("%v", args[0])
	tagName := d.extractTagFromXPath(xpath)

	var results []*XMLElement
	if d.root != nil {
		d.findElements(d.root, tagName, &results)
	}

	var interfaceResults []interface{}
	for _, elem := range results {
		interfaceResults = append(interfaceResults, elem)
	}

	return interfaceResults
}

// createElement creates a new element
// Syntax: CreateElement(tagName)
func (d *MsXML2DOMDocument) createElement(args []interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}

	tagName := fmt.Sprintf("%v", args[0])
	return &XMLElement{
		Name:       tagName,
		Attributes: make(map[string]string),
		Children:   make([]*XMLElement, 0),
	}
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

func (e *XMLElement) SetProperty(name string, value interface{}) {
	switch strings.ToLower(name) {
	case "nodevalue":
		e.Value = fmt.Sprintf("%v", value)
	case "text":
		e.Value = fmt.Sprintf("%v", value)
	}
}

func (e *XMLElement) CallMethod(name string, args []interface{}) interface{} {
	switch strings.ToLower(name) {
	case "appendchild":
		if len(args) > 0 {
			if child, ok := args[0].(*XMLElement); ok {
				e.Children = append(e.Children, child)
				child.Parent = e
				return child
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
			return interfaceResults
		}
	case "item":
		if len(args) > 0 {
			if idx, ok := args[0].(int); ok && idx >= 0 && idx < len(e.Children) {
				return e.Children[idx]
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
			return e.Attributes[key]
		}
	case "removeattribute":
		if len(args) > 0 {
			key := fmt.Sprintf("%v", args[0])
			delete(e.Attributes, key)
		}
	}
	return nil
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

func (d *MsXML2DOMDocument) parseXMLString(xmlStr string) *XMLElement {
	// Simple XML parser - handles basic cases
	xmlStr = strings.TrimSpace(xmlStr)

	// Remove XML declaration if present
	if strings.HasPrefix(xmlStr, "<?xml") {
		endIdx := strings.Index(xmlStr, "?>")
		if endIdx != -1 {
			xmlStr = strings.TrimSpace(xmlStr[endIdx+2:])
		}
	}

	// Find root element
	startIdx := strings.Index(xmlStr, "<")
	if startIdx == -1 {
		return nil
	}

	xmlStr = xmlStr[startIdx:]

	// Parse the root element
	root, _ := d.parseElement(xmlStr)
	return root
}

// parseElement recursively parses an XML element
func (d *MsXML2DOMDocument) parseElement(xmlStr string) (*XMLElement, int) {
	xmlStr = strings.TrimSpace(xmlStr)

	if !strings.HasPrefix(xmlStr, "<") {
		return nil, 0
	}

	// Find end of opening tag
	endTagStart := strings.Index(xmlStr, ">")
	if endTagStart == -1 {
		return nil, 0
	}

	tagContent := xmlStr[1:endTagStart]

	// Check for self-closing tag
	if strings.HasSuffix(tagContent, "/") {
		// Self-closing tag
		tagName := strings.TrimSpace(strings.TrimSuffix(tagContent, "/"))
		elem := &XMLElement{
			Name:       tagName,
			Attributes: make(map[string]string),
			Children:   make([]*XMLElement, 0),
		}
		return elem, endTagStart + 1
	}

	// Parse tag name and attributes
	parts := strings.Fields(tagContent)
	if len(parts) == 0 {
		return nil, 0
	}

	tagName := parts[0]
	elem := &XMLElement{
		Name:       tagName,
		Attributes: make(map[string]string),
		Children:   make([]*XMLElement, 0),
	}

	// Find closing tag
	closingTag := "</" + tagName + ">"
	closeIdx := strings.Index(xmlStr, closingTag)
	if closeIdx == -1 {
		return elem, endTagStart + 1
	}

	// Extract content between opening and closing tags
	contentStart := endTagStart + 1
	contentEnd := closeIdx
	content := xmlStr[contentStart:contentEnd]

	// Parse child elements and text
	d.parseContent(elem, content)

	// Return element and position after closing tag
	return elem, closeIdx + len(closingTag)
}

// parseContent extracts child elements and text from element content
func (d *MsXML2DOMDocument) parseContent(parent *XMLElement, content string) {
	content = strings.TrimSpace(content)
	if content == "" {
		return
	}

	// Parse child elements
	pos := 0
	for pos < len(content) {
		// Skip whitespace and text
		for pos < len(content) && content[pos] != '<' {
			pos++
		}

		if pos >= len(content) {
			break
		}

		// Check if this is a closing tag
		if strings.HasPrefix(content[pos:], "</") {
			break
		}

		// Parse child element
		child, consumed := d.parseElement(content[pos:])
		if child == nil {
			break
		}

		parent.Children = append(parent.Children, child)
		child.Parent = parent
		pos += consumed
	}

	// If no children, the content is text
	if len(parent.Children) == 0 {
		// Remove tags and store as value
		textContent := content
		// Remove any remaining tags
		for strings.Contains(textContent, "<") && strings.Contains(textContent, ">") {
			startIdx := strings.Index(textContent, "<")
			endIdx := strings.Index(textContent, ">")
			if startIdx != -1 && endIdx != -1 && endIdx > startIdx {
				textContent = textContent[:startIdx] + textContent[endIdx+1:]
			} else {
				break
			}
		}
		parent.Value = strings.TrimSpace(textContent)
	}
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

func (d *MsXML2DOMDocument) extractTagFromXPath(xpath string) string {
	// Simple XPath extraction
	parts := strings.Split(xpath, "/")
	for i := len(parts) - 1; i >= 0; i-- {
		part := strings.TrimSpace(parts[i])
		if part != "" && part != "." && !strings.HasPrefix(part, "@") {
			return strings.ToLower(part)
		}
	}
	return ""
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
