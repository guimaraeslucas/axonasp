/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package axonvm

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"net/http"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"

	"golang.org/x/net/html/charset"
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
	ctx             *VM
}

// NewMsXML2ServerXMLHTTP creates a new ServerXMLHTTP instance
func NewMsXML2ServerXMLHTTP(ctx *VM) *MsXML2ServerXMLHTTP {
	return &MsXML2ServerXMLHTTP{
		headers:         make(map[string]string),
		responseHeaders: make(map[string]string),
		readyState:      0,
		timeout:         30 * time.Second,
		async:           false,
		ctx:             ctx,
	}
}

func (s *MsXML2ServerXMLHTTP) legacyGetProperty(name string) interface{} {
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
			return []byte{}
		}
		return s.responseBody
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

func (s *MsXML2ServerXMLHTTP) legacySetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "timeout":
		s.timeout = time.Duration(toInt(value)) * time.Second
	}
	return nil
}

func (s *MsXML2ServerXMLHTTP) legacyCallMethod(name string, args ...interface{}) (interface{}, error) {
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

	// Provide default headers using chrome for safety
	if req.Header.Get("User-Agent") == "" {
		req.Header.Set("User-Agent", "Mozilla/5.0 (Windows NT 11.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36 Edg/145.0.0.0 AxonASPServer/1.0")
	}
	if req.Header.Get("Accept") == "" {
		req.Header.Set("Accept", "*/*")
	}
	if req.Header.Get("Accept-Language") == "" {
		req.Header.Set("Accept-Language", "en-US,en;q=0.9")
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
	contentType := resp.Header.Get("Content-Type")
	s.responseText = decodeResponseText(data, contentType)
	s.status = resp.StatusCode
	s.statusText = resp.Status

	// Store response headers
	for k, v := range resp.Header {
		if len(v) > 0 {
			s.responseHeaders[k] = v[0]
		}
	}

	// Parse XML if response is XML
	if s.isXMLResponse(contentType, s.responseText) {
		doc := NewMsXML2DOMDocument(s.ctx)
		if doc != nil {
			if ok := doc.loadXMLBytes(data, contentType); !ok {
				doc.loadXML([]interface{}{s.responseText})
			}
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
		buf := vbArrayToBytes(v.Values)
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

// ============================================================================
// MsXML2DOMDocument - XML Document Object Model
// ============================================================================

type MsXML2DOMDocument struct {
	xmlContent          string
	root                *XMLElement
	async               bool
	parseError          *ParseError
	serverHTTPRequest   bool
	resolveExternals    bool
	validateOnParse     bool
	preserveWhiteSpace  bool
	selectionLanguage   string
	selectionNamespaces string
	ctx                 *VM
}

// ParseError represents XML parsing errors
type ParseError struct {
	ctx         *VM
	ErrorCode   int
	ErrorReason string
	FilePos     int
	Line        int
	LinePos     int
	SrcText     string
	URL         string
}

// xmlParseErrorDetails keeps parsed XML error position details for ParseError mapping.
type xmlParseErrorDetails struct {
	err     error
	filePos int
	line    int
	linePos int
}

func (e *xmlParseErrorDetails) Error() string {
	if e == nil || e.err == nil {
		return ""
	}
	return e.err.Error()
}

// clearParseError resets ParseError to a success state.
func (d *MsXML2DOMDocument) clearParseError() {
	if d == nil || d.parseError == nil {
		return
	}
	d.parseError.ErrorCode = 0
	d.parseError.ErrorReason = ""
	d.parseError.FilePos = 0
	d.parseError.Line = 0
	d.parseError.LinePos = 0
	d.parseError.SrcText = ""
	d.parseError.URL = ""
}

// setParseErrorFromReason populates ParseError for non-XML-parser failures.
func (d *MsXML2DOMDocument) setParseErrorFromReason(code int, reason string, srcText string, url string) {
	if d == nil || d.parseError == nil {
		return
	}
	d.parseError.ErrorCode = code
	d.parseError.ErrorReason = reason
	d.parseError.FilePos = 0
	d.parseError.Line = 0
	d.parseError.LinePos = 0
	d.parseError.SrcText = srcText
	d.parseError.URL = url
}

// setParseErrorFromXMLError maps XML parser failures to ParseError fields.
func (d *MsXML2DOMDocument) setParseErrorFromXMLError(code int, err error, srcText string, url string) {
	if d == nil || d.parseError == nil {
		return
	}
	reason := "Failed to parse XML"
	filePos := 0
	line := 0
	linePos := 0
	if err != nil {
		reason = err.Error()
		if details, ok := err.(*xmlParseErrorDetails); ok {
			filePos = details.filePos
			line = details.line
			linePos = details.linePos
		}
	}
	d.parseError.ErrorCode = code
	d.parseError.ErrorReason = reason
	d.parseError.FilePos = filePos
	d.parseError.Line = line
	d.parseError.LinePos = linePos
	d.parseError.SrcText = srcText
	d.parseError.URL = url
}

// computeXMLLinePosFromOffset converts a byte offset into 1-based line and column values.
func computeXMLLinePosFromOffset(xmlStr string, offset int) (int, int) {
	if offset < 0 {
		offset = 0
	}
	if offset > len(xmlStr) {
		offset = len(xmlStr)
	}
	line := 1
	linePos := 1
	for i := 0; i < offset; i++ {
		if xmlStr[i] == '\n' {
			line++
			linePos = 1
			continue
		}
		linePos++
	}
	return line, linePos
}

// XMLNodeList represents an MSXML IXMLDOMNodeList
type XMLNodeList struct {
	ctx   *VM
	items []*XMLElement
	next  int
}

func (l *XMLNodeList) legacyGetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "length":
		return len(l.items)
	case "count":
		return len(l.items)
	}
	return nil
}

func (l *XMLNodeList) legacySetProperty(name string, value interface{}) error {
	return nil
}

func (l *XMLNodeList) legacyCallMethod(name string, args ...interface{}) (interface{}, error) {
	method := strings.ToLower(strings.TrimSpace(name))
	if method == "" {
		method = "item"
	}
	switch method {
	case "item":
		if len(args) < 1 {
			return nil, nil
		}
		idx := toInt(args[0])
		if idx < 0 || idx >= len(l.items) {
			return nil, nil
		}
		return l.items[idx], nil
	case "nextnode":
		if l.next < 0 || l.next >= len(l.items) {
			return nil, nil
		}
		item := l.items[l.next]
		l.next++
		return item, nil
	}
	return nil, nil
}

func (l *XMLNodeList) Enumeration() []interface{} {
	items := make([]interface{}, 0, len(l.items))
	for _, item := range l.items {
		items = append(items, item)
	}
	return items
}

// GetName returns the name of the ParseError object

// GetProperty gets a property from the ParseError
func (p *ParseError) legacyGetProperty(name string) interface{} {
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
func (p *ParseError) legacySetProperty(name string, value interface{}) error {
	return nil
}

// CallMethod calls a method on ParseError (none available)
func (p *ParseError) legacyCallMethod(name string, args ...interface{}) (interface{}, error) {
	return nil, nil
}

// XMLElement represents an XML element node
type XMLElement struct {
	ctx        *VM
	Name       string
	LocalName  string
	Namespace  string
	Value      string
	Attributes map[string]string
	AttrNS     map[string]string
	Children   []*XMLElement
	Parent     *XMLElement
}

// NewMsXML2DOMDocument creates a new DOM Document instance
func NewMsXML2DOMDocument(ctx *VM) *MsXML2DOMDocument {
	return &MsXML2DOMDocument{
		async:      false,
		parseError: &ParseError{},
		ctx:        ctx,
	}
}

func (d *MsXML2DOMDocument) legacyGetProperty(name string) interface{} {
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
	case "serverhttprequest":
		return d.serverHTTPRequest
	case "resolveexternals":
		return d.resolveExternals
	case "validateonparse":
		return d.validateOnParse
	case "preservewhitespace":
		return d.preserveWhiteSpace
	case "selectionlanguage":
		return d.selectionLanguage
	case "selectionnamespaces":
		return d.selectionNamespaces
	}
	return nil
}

func (d *MsXML2DOMDocument) legacySetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "async":
		if v, ok := value.(bool); ok {
			d.async = v
		}
	case "serverhttprequest":
		if v, ok := value.(bool); ok {
			d.serverHTTPRequest = v
		}
	case "resolveexternals":
		if v, ok := value.(bool); ok {
			d.resolveExternals = v
		}
	case "validateonparse":
		if v, ok := value.(bool); ok {
			d.validateOnParse = v
		}
	case "preservewhitespace":
		if v, ok := value.(bool); ok {
			d.preserveWhiteSpace = v
		}
	case "selectionlanguage":
		d.selectionLanguage = fmt.Sprintf("%v", value)
	case "selectionnamespaces":
		d.selectionNamespaces = fmt.Sprintf("%v", value)
	}
	return nil
}

func (d *MsXML2DOMDocument) legacyCallMethod(name string, args ...interface{}) (interface{}, error) {
	switch strings.ToLower(name) {
	case "getproperty":
		if len(args) < 1 {
			return nil, nil
		}
		return d.legacyGetProperty(fmt.Sprintf("%v", args[0])), nil
	case "setproperty":
		if len(args) < 2 {
			return nil, nil
		}
		return nil, d.legacySetProperty(fmt.Sprintf("%v", args[0]), args[1])
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
		d.setParseErrorFromReason(-1, "No XML provided", "", "")
		return false
	}

	xmlStr := fmt.Sprintf("%v", args[0])
	d.xmlContent = xmlStr

	root, err := d.parseXMLString(xmlStr)
	if err != nil || root == nil {
		d.setParseErrorFromXMLError(-1, err, xmlStr, "")
		return false
	}

	d.root = root
	d.clearParseError()
	return true
}

// load loads an XML file from URL or path
// Syntax: Load(url)
func (d *MsXML2DOMDocument) load(args []interface{}) interface{} {
	if len(args) < 1 {
		d.setParseErrorFromReason(-1, "No URL provided", "", "")
		return false
	}

	urlStr := fmt.Sprintf("%v", args[0])

	// Try to fetch from URL or path
	var content string

	if strings.HasPrefix(urlStr, "http://") || strings.HasPrefix(urlStr, "https://") {
		client := &http.Client{Timeout: 30 * time.Second}
		req, err := http.NewRequest(http.MethodGet, urlStr, nil)
		if err != nil {
			d.setParseErrorFromReason(-1, err.Error(), "", urlStr)
			return false
		}
		if req.Header.Get("User-Agent") == "" {
			req.Header.Set("User-Agent", "Mozilla/5.0 (Windows NT 11.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36 Edg/145.0.0.0 AxonASP/2.0")
		}
		if req.Header.Get("Accept") == "" {
			req.Header.Set("Accept", "*/*")
		}
		if req.Header.Get("Accept-Language") == "" {
			req.Header.Set("Accept-Language", "en-US,en;q=0.9")
		}

		resp, err := client.Do(req)
		if err != nil {
			d.setParseErrorFromReason(-1, err.Error(), "", urlStr)
			return false
		}
		defer resp.Body.Close()

		if resp.StatusCode >= 400 {
			d.setParseErrorFromReason(resp.StatusCode, resp.Status, "", urlStr)
			return false
		}

		data, err := io.ReadAll(resp.Body)
		if err != nil {
			d.setParseErrorFromReason(-1, err.Error(), "", urlStr)
			return false
		}
		content = decodeResponseText(data, resp.Header.Get("Content-Type"))
	} else {
		if d.ctx != nil {
			fullPath := d.ctx.host.Server().MapPath(urlStr)
			data, errFile := getFileContent(fullPath)
			if errFile != nil {
				d.setParseErrorFromReason(-1, errFile.Error(), "", urlStr)
				return false
			}
			content = data
		} else {
			d.setParseErrorFromReason(-1, "No context available", "", urlStr)
			return false
		}
	}

	d.xmlContent = content
	root, err := d.parseXMLString(content)
	if err != nil || root == nil {
		d.setParseErrorFromXMLError(-1, err, content, urlStr)
		return false
	}

	d.root = root
	d.clearParseError()
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

	fullPath := d.ctx.host.Server().MapPath(filename)
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
		return &XMLNodeList{ctx: d.ctx, items: []*XMLElement{}}
	}

	tagName := strings.ToLower(fmt.Sprintf("%v", args[0]))
	var results []*XMLElement

	if d.root != nil {
		d.findElements(d.root, tagName, &results)
	}

	return &XMLNodeList{ctx: d.ctx, items: results}
}

// selectSingleNode returns the first node that matches the XPath expression.
// Syntax: SelectSingleNode(xpath)
func (d *MsXML2DOMDocument) selectSingleNode(args []interface{}) interface{} {
	nodes := d.selectNodesInternal(args)
	if len(nodes) == 0 {
		return nil
	}
	return nodes[0]
}

// selectNodes returns all nodes that match the XPath expression.
// Syntax: SelectNodes(xpath)
func (d *MsXML2DOMDocument) selectNodes(args []interface{}) interface{} {
	return &XMLNodeList{ctx: d.ctx, items: d.selectNodesInternal(args)}
}

func (d *MsXML2DOMDocument) selectNodesInternal(args []interface{}) []*XMLElement {
	if len(args) < 1 {
		return []*XMLElement{}
	}

	xpathExpr := strings.TrimSpace(fmt.Sprintf("%v", args[0]))
	if xpathExpr == "" {
		return []*XMLElement{}
	}

	if d.root == nil && d.xmlContent != "" {
		if parsed, err := d.parseXMLString(d.xmlContent); err == nil {
			d.root = parsed
		}
	}

	if d.root == nil {
		return []*XMLElement{}
	}

	nsContext := parseSelectionNamespaces(d.selectionNamespaces)

	parts := splitXPathUnion(xpathExpr)
	if len(parts) == 0 {
		parts = []string{xpathExpr}
	}

	all := make([]*XMLElement, 0)
	seen := make(map[*XMLElement]bool)

	for _, part := range parts {
		steps, absolute, err := parseXPathSteps(part)
		if err != nil {
			d.parseError.ErrorCode = -1
			d.parseError.ErrorReason = err.Error()
			continue
		}

		var start *XMLElement
		if absolute {
			start = &XMLElement{Name: "#document", LocalName: "#document", Children: []*XMLElement{d.root}, ctx: d.ctx}
		} else {
			start = d.root
		}

		nodes := evaluateXPath(start, steps, nsContext)
		for _, node := range nodes {
			if node == nil {
				continue
			}
			if !seen[node] {
				seen[node] = true
				all = append(all, node)
			}
		}
	}

	return all
}

type xpathStep struct {
	axis       string
	nodeTest   string
	name       string
	prefix     string
	localName  string
	predicates []string
}

func parseXPathSteps(expr string) ([]xpathStep, bool, error) {
	trimmed := strings.TrimSpace(expr)
	if trimmed == "" {
		return nil, false, fmt.Errorf("invalid XPath expression")
	}

	abs := strings.HasPrefix(trimmed, "/")
	i := 0
	axis := "child"
	if strings.HasPrefix(trimmed, "//") {
		i = 2
		axis = "descendant"
	} else if strings.HasPrefix(trimmed, "/") {
		i = 1
		axis = "child"
	}

	steps := make([]xpathStep, 0)
	for i < len(trimmed) {
		if strings.HasPrefix(trimmed[i:], "//") {
			axis = "descendant"
			i += 2
			continue
		}
		if trimmed[i] == '/' {
			axis = "child"
			i++
			continue
		}

		start := i
		depth := 0
		for i < len(trimmed) {
			if trimmed[i] == '[' {
				depth++
			} else if trimmed[i] == ']' {
				depth--
				if depth < 0 {
					return nil, abs, fmt.Errorf("invalid XPath predicate")
				}
			} else if trimmed[i] == '/' && depth == 0 {
				break
			}
			i++
		}

		segment := strings.TrimSpace(trimmed[start:i])
		if segment == "" {
			continue
		}

		step, err := parseXPathStep(segment, axis)
		if err != nil {
			return nil, abs, err
		}
		steps = append(steps, step)
		axis = "child"
	}

	if len(steps) == 0 {
		return nil, abs, fmt.Errorf("invalid XPath expression")
	}

	return steps, abs, nil
}

func parseXPathStep(segment string, axis string) (xpathStep, error) {
	base := segment
	predicates := make([]string, 0)

	for {
		open := strings.Index(base, "[")
		if open < 0 {
			break
		}
		end := findMatchingBracket(base, open)
		if end < 0 {
			return xpathStep{}, fmt.Errorf("invalid XPath predicate")
		}
		predicates = append(predicates, strings.TrimSpace(base[open+1:end]))
		base = strings.TrimSpace(base[:open] + base[end+1:])
	}

	base = strings.TrimSpace(base)
	step := xpathStep{axis: axis, predicates: predicates, nodeTest: "element", name: base}

	if idx := strings.Index(base, "::"); idx > 0 {
		axisName := strings.ToLower(strings.TrimSpace(base[:idx]))
		if axisName == "child" || axisName == "descendant" || axisName == "self" || axisName == "parent" || axisName == "following-sibling" || axisName == "preceding-sibling" || axisName == "attribute" {
			step.axis = axisName
			base = strings.TrimSpace(base[idx+2:])
			step.name = base
		}
	}

	switch {
	case base == ".":
		step.nodeTest = "self"
		step.name = ""
	case base == "..":
		step.nodeTest = "parent"
		step.name = ""
	case base == "*":
		step.nodeTest = "wildcard"
		step.name = ""
	case strings.EqualFold(base, "text()"):
		step.nodeTest = "text"
		step.name = ""
	case strings.HasPrefix(base, "@"):
		step.nodeTest = "attribute"
		step.name = strings.TrimPrefix(base, "@")
		if step.name == "" {
			return xpathStep{}, fmt.Errorf("invalid attribute XPath")
		}
	default:
		if base == "" {
			return xpathStep{}, fmt.Errorf("invalid XPath step")
		}
	}

	if step.nodeTest == "element" || step.nodeTest == "attribute" {
		step.prefix, step.localName = parseQName(step.name)
	}

	if step.axis == "attribute" {
		step.nodeTest = "attribute"
		if step.name == "" {
			step.name = "*"
			step.localName = "*"
		}
	}

	return step, nil
}

func parseQName(name string) (string, string) {
	name = strings.TrimSpace(name)
	if name == "" {
		return "", ""
	}
	if name == "*" {
		return "", "*"
	}
	if idx := strings.Index(name, ":"); idx >= 0 {
		return name[:idx], name[idx+1:]
	}
	return "", name
}

func findMatchingBracket(s string, open int) int {
	depth := 0
	for i := open; i < len(s); i++ {
		switch s[i] {
		case '[':
			depth++
		case ']':
			depth--
			if depth == 0 {
				return i
			}
		}
	}
	return -1
}

func evaluateXPath(start *XMLElement, steps []xpathStep, nsContext map[string]string) []*XMLElement {
	if start == nil || len(steps) == 0 {
		return []*XMLElement{}
	}

	current := []*XMLElement{start}
	for _, step := range steps {
		next := make([]*XMLElement, 0)
		for _, node := range current {
			if node == nil {
				continue
			}
			next = append(next, applyXPathStep(node, step, nsContext)...)
		}
		current = applyPredicatesToNodes(next, step.predicates, nsContext)
		if len(current) == 0 {
			break
		}
	}

	return current
}

func applyXPathStep(node *XMLElement, step xpathStep, nsContext map[string]string) []*XMLElement {
	results := make([]*XMLElement, 0)

	switch step.axis {
	case "descendant":
		candidates := make([]*XMLElement, 0)
		collectDescendants(node, &candidates)
		for _, candidate := range candidates {
			if nodeMatchesStep(candidate, step, nsContext) {
				results = append(results, projectNodeByStep(candidate, step)...)
			}
		}
	case "following-sibling":
		if node.Parent != nil {
			foundCurrent := false
			for _, sibling := range node.Parent.Children {
				if sibling == node {
					foundCurrent = true
					continue
				}
				if !foundCurrent {
					continue
				}
				if nodeMatchesStep(sibling, step, nsContext) {
					results = append(results, projectNodeByStep(sibling, step)...)
				}
			}
		}
	case "preceding-sibling":
		if node.Parent != nil {
			for _, sibling := range node.Parent.Children {
				if sibling == node {
					break
				}
				if nodeMatchesStep(sibling, step, nsContext) {
					results = append(results, projectNodeByStep(sibling, step)...)
				}
			}
		}
	default:
		if step.nodeTest == "self" {
			results = append(results, node)
			return results
		}
		if step.nodeTest == "parent" {
			if node.Parent != nil {
				results = append(results, node.Parent)
			}
			return results
		}
		if step.nodeTest == "attribute" {
			if step.name == "*" {
				for k, v := range node.Attributes {
					results = append(results, &XMLElement{Name: k, LocalName: k, Value: v, Parent: node, Namespace: node.AttrNS[k], ctx: node.ctx})
				}
				return results
			}
			if attrName, val, ns, ok := lookupAttribute(node, step, nsContext); ok {
				results = append(results, &XMLElement{Name: attrName, LocalName: attrName, Value: val, Parent: node, Namespace: ns, ctx: node.ctx})
			}
			return results
		}

		for _, child := range node.Children {
			if nodeMatchesStep(child, step, nsContext) {
				results = append(results, projectNodeByStep(child, step)...)
			}
		}
	}

	return results
}

func collectDescendants(node *XMLElement, out *[]*XMLElement) {
	for _, child := range node.Children {
		*out = append(*out, child)
		collectDescendants(child, out)
	}
}

func nodeMatchesStep(node *XMLElement, step xpathStep, nsContext map[string]string) bool {
	switch step.nodeTest {
	case "wildcard":
		return true
	case "text":
		return node.Name == "#text"
	case "element":
		return matchXPathName(node, step, nsContext)
	default:
		return false
	}
}

func projectNodeByStep(node *XMLElement, step xpathStep) []*XMLElement {
	if step.nodeTest != "text" {
		return []*XMLElement{node}
	}
	if node.Name == "#text" {
		return []*XMLElement{node}
	}
	results := make([]*XMLElement, 0)
	for _, child := range node.Children {
		if child.Name == "#text" {
			results = append(results, child)
		}
	}
	return results
}

func matchXPathName(node *XMLElement, step xpathStep, nsContext map[string]string) bool {
	if node == nil {
		return false
	}
	nodeLocal := node.Name
	if node.LocalName != "" {
		nodeLocal = node.LocalName
	}

	if step.localName == "" || step.localName == "*" {
		if step.prefix != "" {
			nsURI := nsContext[step.prefix]
			if nsURI == "" {
				return false
			}
			return node.Namespace == nsURI
		}
		return true
	}
	if !strings.EqualFold(nodeLocal, step.localName) {
		if !strings.EqualFold(node.Name, step.name) {
			return false
		}
	}

	if step.prefix != "" {
		nsURI := nsContext[step.prefix]
		if nsURI == "" {
			return false
		}
		return node.Namespace == nsURI
	}

	if defNS, ok := nsContext[""]; ok && defNS != "" {
		return node.Namespace == defNS
	}

	if strings.EqualFold(node.Name, step.name) {
		return true
	}
	return strings.EqualFold(nodeLocal, step.localName)
}

func applyPredicatesToNodes(nodes []*XMLElement, predicates []string, nsContext map[string]string) []*XMLElement {
	filtered := nodes
	for _, predicate := range predicates {
		predicate = strings.TrimSpace(predicate)
		if predicate == "" {
			continue
		}

		if n, ok := parsePredicatePosition(predicate); ok {
			if n <= 0 || n > len(filtered) {
				return []*XMLElement{}
			}
			filtered = []*XMLElement{filtered[n-1]}
			continue
		}

		next := make([]*XMLElement, 0, len(filtered))
		for i, node := range filtered {
			if evaluateNodePredicate(node, predicate, i+1, len(filtered), nsContext) {
				next = append(next, node)
			}
		}
		filtered = next
		if len(filtered) == 0 {
			return filtered
		}
	}
	return filtered
}

func parsePredicatePosition(predicate string) (int, bool) {
	if n, err := strconv.Atoi(strings.TrimSpace(predicate)); err == nil {
		return n, true
	}
	re := regexp.MustCompile(`(?i)^position\(\)\s*=\s*(\d+)$`)
	m := re.FindStringSubmatch(strings.TrimSpace(predicate))
	if len(m) == 2 {
		n, _ := strconv.Atoi(m[1])
		return n, true
	}
	return 0, false
}

func evaluateNodePredicate(node *XMLElement, predicate string, position int, setSize int, nsContext map[string]string) bool {
	predicate = strings.TrimSpace(predicate)
	if predicate == "" {
		return true
	}

	if strings.Contains(strings.ToLower(predicate), " or ") {
		for _, part := range splitTopLevelLogical(predicate, "or") {
			if evaluateNodePredicate(node, part, position, setSize, nsContext) {
				return true
			}
		}
		return false
	}

	if strings.Contains(strings.ToLower(predicate), " and ") {
		for _, part := range splitTopLevelLogical(predicate, "and") {
			if !evaluateNodePredicate(node, part, position, setSize, nsContext) {
				return false
			}
		}
		return true
	}

	if strings.HasPrefix(strings.ToLower(predicate), "not(") && strings.HasSuffix(predicate, ")") {
		inner := strings.TrimSpace(predicate[4 : len(predicate)-1])
		return !evaluateNodePredicate(node, inner, position, setSize, nsContext)
	}

	// Support position() function
	positionFuncRe := regexp.MustCompile(`(?i)^position\(\)\s*(=|!=|<>|>=|<=|>|<)\s*(\d+)$`)
	if m := positionFuncRe.FindStringSubmatch(predicate); len(m) == 3 {
		op := m[1]
		posVal, _ := strconv.Atoi(m[2])
		return compareXPathValues(strconv.Itoa(position), strconv.Itoa(posVal), op)
	}

	// Support boolean functions: true(), false()
	if strings.EqualFold(strings.TrimSpace(predicate), "true()") {
		return true
	}
	if strings.EqualFold(strings.TrimSpace(predicate), "false()") {
		return false
	}

	if strings.EqualFold(predicate, "last()") {
		return position == setSize
	}

	// Support last() with comparison
	lastFuncRe := regexp.MustCompile(`(?i)^last\(\)\s*(=|!=|<>|>=|<=|>|<)\s*(\d+)$`)
	if m := lastFuncRe.FindStringSubmatch(predicate); len(m) == 3 {
		op := m[1]
		posVal, _ := strconv.Atoi(m[2])
		return compareXPathValues(strconv.Itoa(setSize), strconv.Itoa(posVal), op)
	}

	if n, ok := parsePredicatePosition(predicate); ok {
		return position == n
	}

	containsRe := regexp.MustCompile(`(?i)^contains\(\s*(.+?)\s*,\s*(['\"].*['\"])\s*\)$`)
	if m := containsRe.FindStringSubmatch(predicate); len(m) == 3 {
		lhs := evaluateXPathValueExpr(node, m[1], nsContext)
		rhs := trimXPathLiteral(m[2])
		return strings.Contains(lhs, rhs)
	}

	startsWithRe := regexp.MustCompile(`(?i)^starts-with\(\s*(.+?)\s*,\s*(['\"].*['\"])\s*\)$`)
	if m := startsWithRe.FindStringSubmatch(predicate); len(m) == 3 {
		lhs := evaluateXPathValueExpr(node, m[1], nsContext)
		rhs := trimXPathLiteral(m[2])
		return strings.HasPrefix(lhs, rhs)
	}

	// Support substring-after() function
	substringAfterRe := regexp.MustCompile(`(?i)^substring-after\(\s*(.+?)\s*,\s*(['\"].*['\"])\s*\)\s*(=|!=|<>|>=|<=|>|<)?\s*(.*)$`)
	if m := substringAfterRe.FindStringSubmatch(predicate); len(m) >= 3 {
		lhs := evaluateXPathValueExpr(node, m[1], nsContext)
		search := trimXPathLiteral(m[2])
		idx := strings.Index(lhs, search)
		if idx < 0 {
			return false
		}
		result := lhs[idx+len(search):]
		if len(m) >= 5 && m[3] != "" {
			op := m[3]
			rhs := trimXPathLiteral(m[4])
			return compareXPathValues(result, rhs, op)
		}
		return result != ""
	}

	// Support substring-before() function
	substringBeforeRe := regexp.MustCompile(`(?i)^substring-before\(\s*(.+?)\s*,\s*(['\"].*['\"])\s*\)\s*(=|!=|<>|>=|<=|>|<)?\s*(.*)$`)
	if m := substringBeforeRe.FindStringSubmatch(predicate); len(m) >= 3 {
		lhs := evaluateXPathValueExpr(node, m[1], nsContext)
		search := trimXPathLiteral(m[2])
		idx := strings.Index(lhs, search)
		if idx < 0 {
			return false
		}
		result := lhs[:idx]
		if len(m) >= 5 && m[3] != "" {
			op := m[3]
			rhs := trimXPathLiteral(m[4])
			return compareXPathValues(result, rhs, op)
		}
		return result != ""
	}

	// Support string-length() function
	stringLengthRe := regexp.MustCompile(`(?i)^string-length\(\s*(.+?)\s*\)\s*(=|!=|<>|>=|<=|>|<)\s*(\d+)$`)
	if m := stringLengthRe.FindStringSubmatch(predicate); len(m) == 4 {
		lhs := evaluateXPathValueExpr(node, m[1], nsContext)
		op := m[2]
		rhs := m[3]
		return compareXPathValues(strconv.Itoa(len(lhs)), rhs, op)
	}

	normalizeCompareRe := regexp.MustCompile(`(?i)^normalize-space\(\s*text\(\)\s*\)\s*(=|!=|<>|>=|<=|>|<)\s*(.+)$`)
	if m := normalizeCompareRe.FindStringSubmatch(predicate); len(m) == 3 {
		op := m[1]
		rhs := trimXPathLiteral(m[2])
		lhs := normalizeXPathSpace(getNodeText(node))
		return compareXPathValues(lhs, rhs, op)
	}

	attrExistsRe := regexp.MustCompile(`^@([A-Za-z_][\w:\-.]*)$`)
	if m := attrExistsRe.FindStringSubmatch(predicate); len(m) == 2 {
		step := xpathStep{nodeTest: "attribute", name: m[1]}
		step.prefix, step.localName = parseQName(step.name)
		_, _, _, ok := lookupAttribute(node, step, nsContext)
		return ok
	}

	attrCompareRe := regexp.MustCompile(`^@([A-Za-z_][\w:\-.]*)\s*(=|!=|<>|>=|<=|>|<)\s*(.+)$`)
	if m := attrCompareRe.FindStringSubmatch(predicate); len(m) == 4 {
		attrName := m[1]
		op := m[2]
		rhs := trimXPathLiteral(m[3])
		step := xpathStep{nodeTest: "attribute", name: attrName}
		step.prefix, step.localName = parseQName(step.name)
		_, lhs, _, ok := lookupAttribute(node, step, nsContext)
		if !ok {
			return false
		}
		return compareXPathValues(lhs, rhs, op)
	}

	textCompareRe := regexp.MustCompile(`(?i)^text\(\)\s*(=|!=|<>|>=|<=|>|<)\s*(.+)$`)
	if m := textCompareRe.FindStringSubmatch(predicate); len(m) == 3 {
		op := m[1]
		rhs := trimXPathLiteral(m[2])
		lhs := getNodeText(node)
		return compareXPathValues(lhs, rhs, op)
	}

	childCompareRe := regexp.MustCompile(`^([A-Za-z_][\w:\-.]*)\s*(=|!=|<>|>=|<=|>|<)\s*(.+)$`)
	if m := childCompareRe.FindStringSubmatch(predicate); len(m) == 4 {
		childName := m[1]
		op := m[2]
		rhs := trimXPathLiteral(m[3])
		childStep := xpathStep{nodeTest: "element", name: childName}
		childStep.prefix, childStep.localName = parseQName(childName)
		for _, child := range node.Children {
			if matchXPathName(child, childStep, nsContext) {
				if compareXPathValues(getNodeText(child), rhs, op) {
					return true
				}
			}
		}
		return false
	}

	childExistsStep := xpathStep{nodeTest: "element", name: predicate}
	childExistsStep.prefix, childExistsStep.localName = parseQName(predicate)
	if childExistsStep.localName != "" {
		for _, child := range node.Children {
			if matchXPathName(child, childExistsStep, nsContext) {
				return true
			}
		}
	}

	return false
}

func evaluateXPathValueExpr(node *XMLElement, expr string, nsContext map[string]string) string {
	expr = strings.TrimSpace(expr)
	if expr == "" {
		return ""
	}

	if strings.EqualFold(expr, "text()") {
		return getNodeText(node)
	}
	if strings.EqualFold(expr, "normalize-space(text())") {
		return normalizeXPathSpace(getNodeText(node))
	}
	if strings.HasPrefix(expr, "@") {
		step := xpathStep{nodeTest: "attribute", name: strings.TrimPrefix(expr, "@")}
		step.prefix, step.localName = parseQName(step.name)
		_, value, _, ok := lookupAttribute(node, step, nsContext)
		if ok {
			return value
		}
		return ""
	}

	childStep := xpathStep{nodeTest: "element", name: expr}
	childStep.prefix, childStep.localName = parseQName(expr)
	for _, child := range node.Children {
		if matchXPathName(child, childStep, nsContext) {
			return getNodeText(child)
		}
	}

	return expr
}

func compareXPathValues(lhs string, rhs string, op string) bool {
	lf, lErr := strconv.ParseFloat(strings.TrimSpace(lhs), 64)
	rf, rErr := strconv.ParseFloat(strings.TrimSpace(rhs), 64)
	if lErr == nil && rErr == nil {
		switch op {
		case "=":
			return lf == rf
		case "!=", "<>":
			return lf != rf
		case ">":
			return lf > rf
		case "<":
			return lf < rf
		case ">=":
			return lf >= rf
		case "<=":
			return lf <= rf
		}
	}

	cmp := strings.Compare(lhs, rhs)
	switch op {
	case "=":
		return cmp == 0
	case "!=", "<>":
		return cmp != 0
	case ">":
		return cmp > 0
	case "<":
		return cmp < 0
	case ">=":
		return cmp >= 0
	case "<=":
		return cmp <= 0
	default:
		return false
	}
}

func getNodeText(node *XMLElement) string {
	if node == nil {
		return ""
	}
	if node.Name == "#text" {
		return node.Value
	}
	if node.Value != "" {
		return node.Value
	}
	var text strings.Builder
	for _, child := range node.Children {
		if child.Name == "#text" {
			text.WriteString(child.Value)
		}
	}
	return text.String()
}

func lookupAttribute(node *XMLElement, step xpathStep, nsContext map[string]string) (string, string, string, bool) {
	if node == nil {
		return "", "", "", false
	}

	targetLocal := step.localName
	if targetLocal == "" {
		targetLocal = step.name
	}
	targetNS := ""
	if step.prefix != "" {
		targetNS = nsContext[step.prefix]
		if targetNS == "" {
			return "", "", "", false
		}
	}

	if targetLocal == "*" {
		for name, value := range node.Attributes {
			ns := ""
			if node.AttrNS != nil {
				ns = node.AttrNS[name]
			}
			if targetNS == "" || ns == targetNS {
				return name, value, ns, true
			}
		}
		return "", "", "", false
	}

	if v, ok := node.Attributes[targetLocal]; ok {
		ns := ""
		if node.AttrNS != nil {
			ns = node.AttrNS[targetLocal]
		}
		if targetNS == "" || ns == targetNS {
			return targetLocal, v, ns, true
		}
	}

	if idx := strings.Index(targetLocal, ":"); idx >= 0 && idx+1 < len(targetLocal) {
		local := targetLocal[idx+1:]
		if v, ok := node.Attributes[local]; ok {
			ns := ""
			if node.AttrNS != nil {
				ns = node.AttrNS[local]
			}
			if targetNS == "" || ns == targetNS {
				return local, v, ns, true
			}
		}
	}

	return "", "", "", false
}

func trimXPathLiteral(raw string) string {
	raw = strings.TrimSpace(raw)
	if len(raw) >= 2 {
		if (raw[0] == '\'' && raw[len(raw)-1] == '\'') || (raw[0] == '"' && raw[len(raw)-1] == '"') {
			return raw[1 : len(raw)-1]
		}
	}
	return raw
}

func splitTopLevelLogical(expr string, op string) []string {
	parts := make([]string, 0)
	depthParen := 0
	depthBracket := 0
	inSingle := false
	inDouble := false
	last := 0
	lower := strings.ToLower(expr)
	needle := " " + strings.ToLower(op) + " "

	for i := 0; i < len(expr); i++ {
		ch := expr[i]
		switch ch {
		case '\'':
			if !inDouble {
				inSingle = !inSingle
			}
		case '"':
			if !inSingle {
				inDouble = !inDouble
			}
		case '(':
			if !inSingle && !inDouble {
				depthParen++
			}
		case ')':
			if !inSingle && !inDouble && depthParen > 0 {
				depthParen--
			}
		case '[':
			if !inSingle && !inDouble {
				depthBracket++
			}
		case ']':
			if !inSingle && !inDouble && depthBracket > 0 {
				depthBracket--
			}
		}

		if inSingle || inDouble || depthParen > 0 || depthBracket > 0 {
			continue
		}

		if i+len(needle) <= len(expr) && lower[i:i+len(needle)] == needle {
			part := strings.TrimSpace(expr[last:i])
			if part != "" {
				parts = append(parts, part)
			}
			last = i + len(needle)
			i = last - 1
		}
	}

	if tail := strings.TrimSpace(expr[last:]); tail != "" {
		parts = append(parts, tail)
	}

	if len(parts) == 0 {
		return []string{expr}
	}
	return parts
}

func normalizeXPathSpace(s string) string {
	return strings.Join(strings.Fields(strings.TrimSpace(s)), " ")
}

func splitXPathUnion(expr string) []string {
	parts := make([]string, 0)
	depthParen := 0
	depthBracket := 0
	inSingle := false
	inDouble := false
	last := 0

	for i := 0; i < len(expr); i++ {
		ch := expr[i]
		switch ch {
		case '\'':
			if !inDouble {
				inSingle = !inSingle
			}
		case '"':
			if !inSingle {
				inDouble = !inDouble
			}
		case '(':
			if !inSingle && !inDouble {
				depthParen++
			}
		case ')':
			if !inSingle && !inDouble && depthParen > 0 {
				depthParen--
			}
		case '[':
			if !inSingle && !inDouble {
				depthBracket++
			}
		case ']':
			if !inSingle && !inDouble && depthBracket > 0 {
				depthBracket--
			}
		case '|':
			if !inSingle && !inDouble && depthParen == 0 && depthBracket == 0 {
				part := strings.TrimSpace(expr[last:i])
				if part != "" {
					parts = append(parts, part)
				}
				last = i + 1
			}
		}
	}

	if tail := strings.TrimSpace(expr[last:]); tail != "" {
		parts = append(parts, tail)
	}

	if len(parts) == 0 {
		return []string{expr}
	}
	return parts
}

func parseSelectionNamespaces(raw string) map[string]string {
	res := map[string]string{}
	raw = strings.TrimSpace(raw)
	if raw == "" {
		return res
	}
	re := regexp.MustCompile(`(?i)xmlns(?::([A-Za-z_][\w\-.]*))?\s*=\s*['\"]([^'\"]+)['\"]`)
	matches := re.FindAllStringSubmatch(raw, -1)
	for _, m := range matches {
		prefix := ""
		if len(m) > 1 {
			prefix = strings.TrimSpace(m[1])
		}
		uri := ""
		if len(m) > 2 {
			uri = strings.TrimSpace(m[2])
		}
		res[prefix] = uri
	}
	return res
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
		LocalName:  tagName,
		Attributes: make(map[string]string),
		AttrNS:     make(map[string]string),
		Children:   make([]*XMLElement, 0),
		ctx:        d.ctx,
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
		ctx:   d.ctx,
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
		Name:      name,
		LocalName: name,
		ctx:       d.ctx,
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
		if elem.ctx == nil {
			elem.ctx = d.ctx
		}
		if d.root == nil {
			d.root = elem
		} else {
			appendXMLChild(d.root, elem)
		}
		return elem
	}

	return nil
}

// Helper methods for XMLElement (implements Component interface)

func (e *XMLElement) legacyGetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "nodename":
		return e.Name
	case "nodevalue":
		return e.Value
	case "text":
		if e.Name == "#text" {
			return e.Value
		}
		var text strings.Builder
		appendXMLNodeText(&text, e)
		return text.String()
	case "xml":
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

func (e *XMLElement) legacySetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "nodevalue":
		e.Value = fmt.Sprintf("%v", value)
	case "text":
		setXMLNodeText(e, fmt.Sprintf("%v", value))
	}
	return nil
}

func (e *XMLElement) legacyCallMethod(name string, args ...interface{}) (interface{}, error) {
	switch strings.ToLower(name) {
	case "appendchild":
		if len(args) > 0 {
			if child, ok := args[0].(*XMLElement); ok {
				appendXMLChild(e, child)
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
			return interfaceResults, nil
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
	case "selectsinglenode":
		if len(args) < 1 {
			return nil, nil
		}
		xpathExpr := strings.TrimSpace(fmt.Sprintf("%v", args[0]))
		if xpathExpr == "" {
			return nil, nil
		}
		steps, absolute, err := parseXPathSteps(xpathExpr)
		if err != nil {
			return nil, nil
		}
		start := e
		if absolute {
			root := e
			for root.Parent != nil {
				root = root.Parent
			}
			start = &XMLElement{Name: "#document", Children: []*XMLElement{root}, ctx: e.ctx}
		}
		nodes := evaluateXPath(start, steps, map[string]string{})
		if len(nodes) == 0 {
			return nil, nil
		}
		return nodes[0], nil
	case "selectnodes":
		if len(args) < 1 {
			return &XMLNodeList{ctx: e.ctx, items: []*XMLElement{}}, nil
		}
		xpathExpr := strings.TrimSpace(fmt.Sprintf("%v", args[0]))
		if xpathExpr == "" {
			return &XMLNodeList{ctx: e.ctx, items: []*XMLElement{}}, nil
		}
		steps, absolute, err := parseXPathSteps(xpathExpr)
		if err != nil {
			return &XMLNodeList{ctx: e.ctx, items: []*XMLElement{}}, nil
		}
		start := e
		if absolute {
			root := e
			for root.Parent != nil {
				root = root.Parent
			}
			start = &XMLElement{Name: "#document", Children: []*XMLElement{root}, ctx: e.ctx}
		}
		nodes := evaluateXPath(start, steps, map[string]string{})
		return &XMLNodeList{ctx: e.ctx, items: nodes}, nil
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
	var buf bytes.Buffer
	writeXMLNode(&buf, e)
	return buf.String()
}

// appendXMLChild attaches one XML child node to one parent and propagates the VM context.
func appendXMLChild(parent *XMLElement, child *XMLElement) {
	if parent == nil || child == nil {
		return
	}
	if child.ctx == nil {
		child.ctx = parent.ctx
	}
	child.Parent = parent
	parent.Children = append(parent.Children, child)
	if child.Name == "#text" {
		parent.Value = ""
	}
}

// appendXMLNodeText collects one node's text content using descendant text-node traversal.
func appendXMLNodeText(buf *strings.Builder, node *XMLElement) {
	if buf == nil || node == nil {
		return
	}
	if node.Name == "#text" {
		buf.WriteString(node.Value)
		return
	}
	if len(node.Children) == 0 && node.Value != "" {
		buf.WriteString(node.Value)
		return
	}
	for _, child := range node.Children {
		appendXMLNodeText(buf, child)
	}
}

// setXMLNodeText replaces one node's content with one text child while keeping text-node assignments direct.
func setXMLNodeText(node *XMLElement, text string) {
	if node == nil {
		return
	}
	if node.Name == "#text" {
		node.Value = text
		return
	}
	node.Value = ""
	node.Children = node.Children[:0]
	if text == "" {
		return
	}
	textNode := &XMLElement{Name: "#text", Value: text, Parent: node, ctx: node.ctx}
	node.Children = append(node.Children, textNode)
}

// writeXMLNode serializes one XML node without adding pretty-print whitespace.
func writeXMLNode(buf *bytes.Buffer, node *XMLElement) {
	if buf == nil || node == nil {
		return
	}
	if node.Name == "#text" {
		buf.WriteString(node.Value)
		return
	}
	buf.WriteByte('<')
	buf.WriteString(node.Name)
	for k, v := range node.Attributes {
		buf.WriteByte(' ')
		buf.WriteString(k)
		buf.WriteString(`="`)
		buf.WriteString(v)
		buf.WriteByte('"')
	}
	if len(node.Children) == 0 && node.Value == "" {
		buf.WriteString("/>")
		return
	}
	buf.WriteByte('>')
	if len(node.Children) == 0 {
		buf.WriteString(node.Value)
	} else {
		for _, child := range node.Children {
			writeXMLNode(buf, child)
		}
	}
	buf.WriteString("</")
	buf.WriteString(node.Name)
	buf.WriteByte('>')
}

// Private helper methods for MsXML2DOMDocument

func (d *MsXML2DOMDocument) parseXMLString(xmlStr string) (*XMLElement, error) {
	if strings.TrimSpace(xmlStr) == "" {
		return nil, &xmlParseErrorDetails{err: fmt.Errorf("empty xml"), filePos: 0, line: 0, linePos: 0}
	}

	decoder := xml.NewDecoder(strings.NewReader(xmlStr))
	decoder.Strict = true
	decoder.CharsetReader = charset.NewReaderLabel

	var root *XMLElement
	stack := make([]*XMLElement, 0)

	for {
		tok, err := decoder.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			filePos := int(decoder.InputOffset())
			line, linePos := computeXMLLinePosFromOffset(xmlStr, filePos)
			return nil, &xmlParseErrorDetails{
				err:     err,
				filePos: filePos,
				line:    line,
				linePos: linePos,
			}
		}

		switch t := tok.(type) {
		case xml.StartElement:
			name := t.Name.Local
			if t.Name.Space != "" {
				name = t.Name.Local
			}
			node := &XMLElement{
				Name:       name,
				LocalName:  t.Name.Local,
				Namespace:  t.Name.Space,
				Attributes: make(map[string]string),
				AttrNS:     make(map[string]string),
				Children:   make([]*XMLElement, 0),
				ctx:        d.ctx,
			}
			for _, attr := range t.Attr {
				node.Attributes[attr.Name.Local] = attr.Value
				node.AttrNS[attr.Name.Local] = attr.Name.Space
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
			if !d.preserveWhiteSpace && strings.TrimSpace(text) == "" {
				continue
			}
			if len(stack) == 0 {
				continue
			}
			parent := stack[len(stack)-1]
			textNode := &XMLElement{
				Name:  "#text",
				Value: text,
				ctx:   d.ctx,
			}
			parent.Children = append(parent.Children, textNode)
			textNode.Parent = parent
		}
	}

	if len(stack) != 0 {
		filePos := len(xmlStr)
		line, linePos := computeXMLLinePosFromOffset(xmlStr, filePos)
		return nil, &xmlParseErrorDetails{
			err:     fmt.Errorf("unexpected end of XML document"),
			filePos: filePos,
			line:    line,
			linePos: linePos,
		}
	}

	return root, nil
}

func (d *MsXML2DOMDocument) loadXMLBytes(data []byte, contentType string) bool {
	if len(data) == 0 {
		d.setParseErrorFromReason(-1, "Empty XML", "", "")
		return false
	}

	decoded := decodeResponseText(data, contentType)
	d.xmlContent = decoded
	root, err := d.parseXMLString(decoded)
	if err != nil || root == nil {
		d.setParseErrorFromXMLError(-1, err, decoded, "")
		return false
	}

	d.root = root
	d.clearParseError()
	return true
}

func parseCharsetFromContentType(contentType string) string {
	if contentType == "" {
		return ""
	}
	parts := strings.Split(contentType, ";")
	for _, part := range parts[1:] {
		p := strings.TrimSpace(part)
		if strings.HasPrefix(strings.ToLower(p), "charset=") {
			return strings.Trim(strings.TrimSpace(p[len("charset="):]), "\"")
		}
	}
	return ""
}

func parseXMLDeclEncoding(data []byte) string {
	if len(data) == 0 {
		return ""
	}
	limit := len(data)
	if limit > 512 {
		limit = 512
	}
	chunk := string(data[:limit])
	re := regexp.MustCompile(`(?i)encoding\s*=\s*['\"]([^'\"]+)['\"]`)
	match := re.FindStringSubmatch(chunk)
	if len(match) >= 2 {
		return strings.TrimSpace(match[1])
	}
	return ""
}

func decodeBytesWithCharset(data []byte, charsetName string) ([]byte, error) {
	if charsetName == "" {
		return data, nil
	}
	r, err := charset.NewReaderLabel(strings.ToLower(charsetName), bytes.NewReader(data))
	if err != nil {
		return nil, err
	}
	decoded, err := io.ReadAll(r)
	if err != nil {
		return nil, err
	}
	return decoded, nil
}

func decodeResponseText(data []byte, contentType string) string {
	if len(data) == 0 {
		return ""
	}

	if charsetName := parseCharsetFromContentType(contentType); charsetName != "" {
		if decoded, err := decodeBytesWithCharset(data, charsetName); err == nil {
			return string(decoded)
		}
	}

	if enc := parseXMLDeclEncoding(data); enc != "" {
		if decoded, err := decodeBytesWithCharset(data, enc); err == nil {
			return string(decoded)
		}
	}

	if len(data) >= 3 && bytes.Equal(data[:3], []byte{0xEF, 0xBB, 0xBF}) {
		data = data[3:]
	}

	if len(data) >= 2 {
		switch {
		case data[0] == 0xFF && data[1] == 0xFE:
			if decoded, err := decodeBytesWithCharset(data, "utf-16le"); err == nil {
				return string(decoded)
			}
		case data[0] == 0xFE && data[1] == 0xFF:
			if decoded, err := decodeBytesWithCharset(data, "utf-16be"); err == nil {
				return string(decoded)
			}
		}
	}

	return string(data)
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
	writeXMLNode(&buf, elem)
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

func (x *MsXML2ServerXMLHTTP) DispatchPropertyGet(name string) Value {
	return legacyInterfaceToValue(x.legacyGetProperty(name), x.ctx)
}
func (x *MsXML2ServerXMLHTTP) DispatchPropertySet(name string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	x.legacySetProperty(name, legacyValueToInterface(args[0], x.ctx))
	return true
}

func (x *MsXML2ServerXMLHTTP) DispatchMethod(name string, args []Value) Value {
	var iArgs []interface{}
	for _, a := range args {
		iArgs = append(iArgs, legacyValueToInterface(a, x.ctx))
	}
	res, _ := x.legacyCallMethod(name, iArgs...)
	return legacyInterfaceToValue(res, x.ctx)
}

func (x *MsXML2DOMDocument) DispatchPropertyGet(name string) Value {
	return legacyInterfaceToValue(x.legacyGetProperty(name), x.ctx)
}
func (x *MsXML2DOMDocument) DispatchPropertySet(name string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	x.legacySetProperty(name, legacyValueToInterface(args[0], x.ctx))
	return true
}

func (x *MsXML2DOMDocument) DispatchMethod(name string, args []Value) Value {
	var iArgs []interface{}
	for _, a := range args {
		iArgs = append(iArgs, legacyValueToInterface(a, x.ctx))
	}
	res, _ := x.legacyCallMethod(name, iArgs...)
	return legacyInterfaceToValue(res, x.ctx)
}

func (x *XMLNodeList) DispatchPropertyGet(name string) Value {
	return legacyInterfaceToValue(x.legacyGetProperty(name), x.ctx)
}
func (x *XMLNodeList) DispatchPropertySet(name string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	x.legacySetProperty(name, legacyValueToInterface(args[0], x.ctx))
	return true
}

func (x *XMLNodeList) DispatchMethod(name string, args []Value) Value {
	var iArgs []interface{}
	for _, a := range args {
		iArgs = append(iArgs, legacyValueToInterface(a, x.ctx))
	}
	res, _ := x.legacyCallMethod(name, iArgs...)
	return legacyInterfaceToValue(res, x.ctx)
}

func (x *ParseError) DispatchPropertyGet(name string) Value {
	return legacyInterfaceToValue(x.legacyGetProperty(name), x.ctx)
}
func (x *ParseError) DispatchPropertySet(name string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	x.legacySetProperty(name, legacyValueToInterface(args[0], x.ctx))
	return true
}

func (x *ParseError) DispatchMethod(name string, args []Value) Value {
	var iArgs []interface{}
	for _, a := range args {
		iArgs = append(iArgs, legacyValueToInterface(a, x.ctx))
	}
	res, _ := x.legacyCallMethod(name, iArgs...)
	return legacyInterfaceToValue(res, x.ctx)
}

func (x *XMLElement) DispatchPropertyGet(name string) Value {
	return legacyInterfaceToValue(x.legacyGetProperty(name), x.ctx)
}
func (x *XMLElement) DispatchPropertySet(name string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	x.legacySetProperty(name, legacyValueToInterface(args[0], x.ctx))
	return true
}

func (x *XMLElement) DispatchMethod(name string, args []Value) Value {
	var iArgs []interface{}
	for _, a := range args {
		iArgs = append(iArgs, legacyValueToInterface(a, x.ctx))
	}
	res, _ := x.legacyCallMethod(name, iArgs...)
	return legacyInterfaceToValue(res, x.ctx)
}

func vbArrayToBytes(arr []Value) []byte {
	buf := make([]byte, len(arr))
	for i, val := range arr {
		buf[i] = byte(toInt(val))
	}
	return buf
}
