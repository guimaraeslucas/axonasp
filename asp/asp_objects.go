package asp

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
		if len(args) > 0 {
			// Implementação simplificada
			return args[0], nil
		}
		return "", nil
	case "URLEncode":
		if len(args) > 0 {
			return urlEncode(args[0].(string)), nil
		}
		return "", nil
	case "HTMLEncode":
		if len(args) > 0 {
			return htmlEncode(args[0].(string)), nil
		}
		return "", nil
	default:
		return nil, nil
	}
}

// RequestObject representa o objeto Request do ASP
type RequestObject struct {
	properties map[string]interface{}
	queryString map[string]interface{}
	form map[string]interface{}
	cookies map[string]interface{}
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
	if val, exists := r.properties[name]; exists {
		return val
	}
	return nil
}

// SetProperty define uma propriedade
func (r *ResponseObject) SetProperty(name string, value interface{}) error {
	r.properties[name] = value
	return nil
}

// CallMethod chama um método do objeto
func (r *ResponseObject) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch name {
	case "Write":
		if len(args) > 0 {
			r.buffer += args[0].(string)
		}
		return nil, nil
	case "Redirect":
		if len(args) > 0 {
			// Simulação de redirect
			r.properties["__redirect__"] = args[0]
		}
		return nil, nil
	case "AddHeader":
		if len(args) >= 2 {
			r.headers[args[0].(string)] = args[1].(string)
		}
		return nil, nil
	default:
		return nil, nil
	}
}

// GetBuffer retorna o conteúdo do buffer de saída
func (r *ResponseObject) GetBuffer() string {
	return r.buffer
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

func urlEncode(s string) string {
	// Implementação simplificada
	replacements := map[string]string{
		" ": "+",
		"&": "%26",
		"=": "%3D",
		"?": "%3F",
	}
	result := s
	for old, new := range replacements {
		result = replaceAll(result, old, new)
	}
	return result
}

func htmlEncode(s string) string {
	replacements := map[string]string{
		"&": "&amp;",
		"<": "&lt;",
		">": "&gt;",
		"\"": "&quot;",
		"'": "&#39;",
	}
	result := s
	for old, new := range replacements {
		result = replaceAll(result, old, new)
	}
	return result
}

func replaceAll(str, old, new string) string {
	result := ""
	for {
		idx := indexOf(str, old)
		if idx == -1 {
			result += str
			break
		}
		result += str[:idx] + new
		str = str[idx+len(old):]
	}
	return result
}

func indexOf(s, substr string) int {
	for i := 0; i <= len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return i
		}
	}
	return -1
}
