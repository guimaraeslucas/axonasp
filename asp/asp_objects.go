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
package asp

import (
	"fmt"
	"html"
	"net/http"
	"net/url"
	"path/filepath"
	"strconv"
	"strings"
)

// ASPObject representa um objeto ASP
type ASPObject interface {
	GetName() string
	GetProperty(name string) interface{}
	SetProperty(name string, value interface{}) error
	CallMethod(name string, args ...interface{}) (interface{}, error)
}

// ServerObject representa o objeto Server do ASP
type ServerObject struct {
	properties map[string]interface{}
}

// NewServerObject cria um novo objeto Server
func NewServerObject() *ServerObject {
	return &ServerObject{
		properties: make(map[string]interface{}),
	}
}

// GetName retorna o nome do objeto
func (s *ServerObject) GetName() string {
	return "Server"
}

// GetProperty obtém uma propriedade
func (s *ServerObject) GetProperty(name string) interface{} {
	if val, exists := s.properties[name]; exists {
		return val
	}
	return nil
}

// SetProperty define uma propriedade
func (s *ServerObject) SetProperty(name string, value interface{}) error {
	s.properties[name] = value
	return nil
}

// CallMethod chama um método do objeto
func (s *ServerObject) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch name {
	case "MapPath":
		return s.mapPath(args), nil
	case "URLEncode":
		return s.urlEncode(args), nil
	case "HTMLEncode":
		return s.htmlEncode(args), nil
	case "GetLastError":
		return s.getLastError(), nil
	case "IsClientConnected":
		return s.isClientConnected(), nil
	case "CreateObject":
		return s.createObject(args)
	default:
		return nil, nil
	}
}

// mapPath converts a virtual path to an absolute file system path
func (s *ServerObject) mapPath(args []interface{}) interface{} {
	if len(args) == 0 {
		return ""
	}

	path, ok := args[0].(string)
	if !ok {
		return ""
	}

	// Get root directory from properties
	rootDir, _ := s.properties["_rootDir"].(string)
	if rootDir == "" {
		rootDir = "./www"
	}

	// Handle different path formats
	if path == "/" || path == "" {
		return rootDir
	}

	// Remove leading slash if present
	if path[0] == '/' || path[0] == '\\' {
		path = path[1:]
	}

	// Join with root directory
	fullPath := filepath.Join(rootDir, path)

	// Convert to absolute path
	absPath, err := filepath.Abs(fullPath)
	if err != nil {
		return fullPath
	}

	return absPath
}

// urlEncode encodes a string for use in URLs (RFC 3986)
func (s *ServerObject) urlEncode(args []interface{}) interface{} {
	if len(args) == 0 {
		return ""
	}

	str, ok := args[0].(string)
	if !ok {
		return ""
	}

	// Use net/url QueryEscape which follows RFC 3986 standard
	return url.QueryEscape(str)
}

// htmlEncode encodes a string for safe HTML output
func (s *ServerObject) htmlEncode(args []interface{}) interface{} {
	if len(args) == 0 {
		return ""
	}

	str, ok := args[0].(string)
	if !ok {
		return ""
	}

	// Use html.EscapeString from standard library
	return html.EscapeString(str)
}

// getLastError returns the last error object (if any)
func (s *ServerObject) getLastError() interface{} {
	// Return nil if no error stored in properties
	if err, exists := s.properties["_lastError"]; exists {
		return err
	}
	return nil
}

// isClientConnected checks if HTTP client is still connected
func (s *ServerObject) isClientConnected() interface{} {
	// Get HTTP request context from properties
	if req, exists := s.properties["_httpRequest"].(*http.Request); exists {
		// Check if request context has been cancelled
		select {
		case <-req.Context().Done():
			return false
		default:
			return true
		}
	}
	// Default to true if no request context available
	return true
}

// createObject creates an ASP COM object (Server.CreateObject)
func (s *ServerObject) createObject(args []interface{}) (interface{}, error) {
	if len(args) == 0 {
		return nil, fmt.Errorf("CreateObject requires an object type")
	}

	objType := fmt.Sprintf("%v", args[0])

	// Get the executor from properties (set by server/executor.go)
	if executor, exists := s.properties["_executor"]; exists {
		// Call the executor's CreateObject method
		// The executor is an interface with CreateObject(string) (interface{}, error)
		if ex, ok := executor.(interface {
			CreateObject(string) (interface{}, error)
		}); ok {
			return ex.CreateObject(objType)
		}
	}

	return nil, fmt.Errorf("CreateObject not available")
}

// RequestObject representa o objeto Request do ASP
type RequestObject struct {
	properties      map[string]interface{}
	queryString     map[string]interface{}
	form            map[string]interface{}
	cookies         map[string]interface{}
	serverVariables map[string]interface{}
}

// NewRequestObject cria um novo objeto Request
func NewRequestObject() *RequestObject {
	return &RequestObject{
		properties:      make(map[string]interface{}),
		queryString:     make(map[string]interface{}),
		form:            make(map[string]interface{}),
		cookies:         make(map[string]interface{}),
		serverVariables: make(map[string]interface{}),
	}
}

// GetName retorna o nome do objeto
func (r *RequestObject) GetName() string {
	return "Request"
}

// GetProperty obtém uma propriedade
func (r *RequestObject) GetProperty(name string) interface{} {
	if val, exists := r.properties[name]; exists {
		return val
	}
	return nil
}

// SetProperty define uma propriedade
func (r *RequestObject) SetProperty(name string, value interface{}) error {
	r.properties[name] = value
	return nil
}

// CallMethod chama um método do objeto
func (r *RequestObject) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch name {
	case "QueryString":
		if len(args) > 0 {
			key := args[0].(string)
			return r.queryString[key], nil
		}
		return r.queryString, nil
	case "Form":
		if len(args) > 0 {
			key := args[0].(string)
			return r.form[key], nil
		}
		return r.form, nil
	case "Cookies":
		if len(args) > 0 {
			key := args[0].(string)
			return r.cookies[key], nil
		}
		return r.cookies, nil
	case "ServerVariables":
		if len(args) > 0 {
			key := args[0].(string)
			return r.serverVariables[key], nil
		}
		return r.serverVariables, nil
	default:
		return nil, nil
	}
}

// ResponseObject representa o objeto Response do ASP
type ResponseObject struct {
	properties map[string]interface{}
	buffer     string
	headers    map[string]string
}

// NewResponseObject cria um novo objeto Response
func NewResponseObject() *ResponseObject {
	return &ResponseObject{
		properties: make(map[string]interface{}),
		headers:    make(map[string]string),
	}
}

// GetName retorna o nome do objeto
func (r *ResponseObject) GetName() string {
	return "Response"
}

// GetProperty obtém uma propriedade
func (r *ResponseObject) GetProperty(name string) interface{} {
	nameLower := strings.ToLower(name)
	if val, exists := r.properties[nameLower]; exists {
		return val
	}
	return nil
}

// SetProperty define uma propriedade
func (r *ResponseObject) SetProperty(name string, value interface{}) error {
	nameLower := strings.ToLower(name)
	r.properties[nameLower] = value
	return nil
}

// CallMethod chama um método do objeto
func (r *ResponseObject) CallMethod(name string, args ...interface{}) (interface{}, error) {
	// Convert method name to lowercase for case-insensitive comparison
	nameLower := strings.ToLower(name)
	switch nameLower {
	case "write":
		return r.write(args), nil
	case "redirect":
		if len(args) > 0 {
			// Simulação de redirect
			r.properties["__redirect__"] = args[0]
		}
		return nil, nil
	case "addheader":
		if len(args) >= 2 {
			r.headers[args[0].(string)] = args[1].(string)
		}
		return nil, nil
	default:
		return nil, nil
	}
}

// write escreve conteúdo no buffer de resposta
// Suporta múltiplos argumentos e conversão automática de tipos
func (r *ResponseObject) write(args []interface{}) interface{} {
	if len(args) == 0 {
		return nil
	}

	// Converter todos os argumentos para string e concatenar
	for _, arg := range args {
		r.buffer += r.toString(arg)
	}

	return nil
}

// toString converte um valor para string seguindo as regras ASP
func (r *ResponseObject) toString(value interface{}) string {
	if value == nil {
		return ""
	}

	switch v := value.(type) {
	case string:
		return v
	case int:
		return fmt.Sprintf("%d", v)
	case int32:
		return fmt.Sprintf("%d", v)
	case int64:
		return fmt.Sprintf("%d", v)
	case float64:
		// ASP converte floats com até 15 dígitos significativos
		s := strconv.FormatFloat(v, 'g', -1, 64)
		return s
	case float32:
		return strconv.FormatFloat(float64(v), 'g', -1, 32)
	case bool:
		if v {
			return "True"
		}
		return "False"
	case []byte:
		return string(v)
	default:
		// Fallback para qualquer outro tipo
		return fmt.Sprintf("%v", v)
	}
}

// GetBuffer retorna o conteúdo do buffer de saída
func (r *ResponseObject) GetBuffer() string {
	return r.buffer
}

// GetContentType retorna o Content-Type definido ou o padrão
func (r *ResponseObject) GetContentType() string {
	if val, exists := r.properties["contenttype"]; exists {
		return fmt.Sprintf("%v", val)
	}
	return "text/html; charset=utf-8"
}

// SessionObject representa o objeto Session do ASP
type SessionObject struct {
	properties map[string]interface{}
}

// NewSessionObject cria um novo objeto Session
func NewSessionObject() *SessionObject {
	return &SessionObject{
		properties: make(map[string]interface{}),
	}
}

// GetName retorna o nome do objeto
func (s *SessionObject) GetName() string {
	return "Session"
}

// GetProperty obtém uma propriedade
func (s *SessionObject) GetProperty(name string) interface{} {
	if val, exists := s.properties[name]; exists {
		return val
	}
	return nil
}

// SetProperty define uma propriedade
func (s *SessionObject) SetProperty(name string, value interface{}) error {
	s.properties[name] = value
	return nil
}

// CallMethod chama um método do objeto
func (s *SessionObject) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return nil, nil
}

// ApplicationObject representa o objeto Application do ASP
type ApplicationObject struct {
	properties map[string]interface{}
}

// NewApplicationObject cria um novo objeto Application
func NewApplicationObject() *ApplicationObject {
	return &ApplicationObject{
		properties: make(map[string]interface{}),
	}
}

// GetName retorna o nome do objeto
func (a *ApplicationObject) GetName() string {
	return "Application"
}

// GetProperty obtém uma propriedade
func (a *ApplicationObject) GetProperty(name string) interface{} {
	if val, exists := a.properties[name]; exists {
		return val
	}
	return nil
}

// SetProperty define uma propriedade
func (a *ApplicationObject) SetProperty(name string, value interface{}) error {
	a.properties[name] = value
	return nil
}

// CallMethod chama um método do objeto
func (a *ApplicationObject) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return nil, nil
}

// ASPContext representa o contexto de execução do ASP
type ASPContext struct {
	Server      *ServerObject
	Request     *RequestObject
	Response    *ResponseObject
	Session     *SessionObject
	Application *ApplicationObject
	Variables   map[string]interface{}
}

// NewASPContext cria um novo contexto ASP
func NewASPContext() *ASPContext {
	return &ASPContext{
		Server:      NewServerObject(),
		Request:     NewRequestObject(),
		Response:    NewResponseObject(),
		Session:     NewSessionObject(),
		Application: NewApplicationObject(),
		Variables:   make(map[string]interface{}),
	}
}

// Funções auxiliares para encode
