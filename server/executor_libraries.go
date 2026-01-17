package server

// Library interface for ASP libraries
type ASPLibrary interface {
	CallMethod(name string, args ...interface{}) (interface{}, error)
	GetProperty(name string) interface{}
	SetProperty(name string, value interface{}) error
}

// JSONLibrary wraps G3JSON for ASPLibrary interface compatibility
type JSONLibrary struct {
	lib *G3JSON
}

// NewJSONLibrary creates a new JSON library instance
func NewJSONLibrary(ctx *ExecutionContext) *JSONLibrary {
	return &JSONLibrary{
		lib: &G3JSON{},
	}
}

// CallMethod calls a method on the JSON library
func (jl *JSONLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return jl.lib.CallMethod(name, args), nil
}

// GetProperty gets a property from the JSON library
func (jl *JSONLibrary) GetProperty(name string) interface{} {
	return jl.lib.GetProperty(name)
}

// SetProperty sets a property on the JSON library
func (jl *JSONLibrary) SetProperty(name string, value interface{}) error {
	jl.lib.SetProperty(name, value)
	return nil
}

// FileSystemLibrary wraps G3FILES for ASPLibrary interface compatibility
type FileSystemLibrary struct {
	lib *G3FILES
}

// NewFileSystemLibrary creates a new FileSystem library instance
func NewFileSystemLibrary(ctx *ExecutionContext) *FileSystemLibrary {
	return &FileSystemLibrary{
		lib: &G3FILES{ctx: ctx},
	}
}

// CallMethod calls a method on the FileSystem library
func (fs *FileSystemLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return fs.lib.CallMethod(name, args), nil
}

// GetProperty gets a property from the FileSystem library
func (fs *FileSystemLibrary) GetProperty(name string) interface{} {
	return fs.lib.GetProperty(name)
}

// SetProperty sets a property on the FileSystem library
func (fs *FileSystemLibrary) SetProperty(name string, value interface{}) error {
	fs.lib.SetProperty(name, value)
	return nil
}

// HTTPLibrary wraps G3HTTP for ASPLibrary interface compatibility
type HTTPLibrary struct {
	lib *G3HTTP
}

// NewHTTPLibrary creates a new HTTP library instance
func NewHTTPLibrary(ctx *ExecutionContext) *HTTPLibrary {
	return &HTTPLibrary{
		lib: &G3HTTP{ctx: ctx},
	}
}

// CallMethod calls a method on the HTTP library
func (hl *HTTPLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return hl.lib.CallMethod(name, args), nil
}

// GetProperty gets a property from the HTTP library
func (hl *HTTPLibrary) GetProperty(name string) interface{} {
	return hl.lib.GetProperty(name)
}

// SetProperty sets a property on the HTTP library
func (hl *HTTPLibrary) SetProperty(name string, value interface{}) error {
	hl.lib.SetProperty(name, value)
	return nil
}

// TemplateLibrary wraps G3TEMPLATE for ASPLibrary interface compatibility
type TemplateLibrary struct {
	lib *G3TEMPLATE
}

// NewTemplateLibrary creates a new Template library instance
func NewTemplateLibrary(ctx *ExecutionContext) *TemplateLibrary {
	return &TemplateLibrary{
		lib: &G3TEMPLATE{ctx: ctx},
	}
}

// CallMethod calls a method on the Template library
func (tl *TemplateLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return tl.lib.CallMethod(name, args), nil
}

// GetProperty gets a property from the Template library
func (tl *TemplateLibrary) GetProperty(name string) interface{} {
	return tl.lib.GetProperty(name)
}

// SetProperty sets a property on the Template library
func (tl *TemplateLibrary) SetProperty(name string, value interface{}) error {
	tl.lib.SetProperty(name, value)
	return nil
}

// MailLibrary wraps G3MAIL for ASPLibrary interface compatibility
type MailLibrary struct {
	lib *G3MAIL
}

// NewMailLibrary creates a new Mail library instance
func NewMailLibrary(ctx *ExecutionContext) *MailLibrary {
	return &MailLibrary{
		lib: &G3MAIL{ctx: ctx},
	}
}

// CallMethod calls a method on the Mail library
func (ml *MailLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return ml.lib.CallMethod(name, args), nil
}

// GetProperty gets a property from the Mail library
func (ml *MailLibrary) GetProperty(name string) interface{} {
	return ml.lib.GetProperty(name)
}

// SetProperty sets a property on the Mail library
func (ml *MailLibrary) SetProperty(name string, value interface{}) error {
	ml.lib.SetProperty(name, value)
	return nil
}

// CryptoLibrary wraps G3CRYPTO for ASPLibrary interface compatibility
type CryptoLibrary struct {
	lib *G3CRYPTO
}

// NewCryptoLibrary creates a new Crypto library instance
func NewCryptoLibrary(ctx *ExecutionContext) *CryptoLibrary {
	return &CryptoLibrary{
		lib: &G3CRYPTO{ctx: ctx},
	}
}

// CallMethod calls a method on the Crypto library
func (cl *CryptoLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return cl.lib.CallMethod(name, args), nil
}

// GetProperty gets a property from the Crypto library
func (cl *CryptoLibrary) GetProperty(name string) interface{} {
	return cl.lib.GetProperty(name)
}

// SetProperty sets a property on the Crypto library
func (cl *CryptoLibrary) SetProperty(name string, value interface{}) error {
	cl.lib.SetProperty(name, value)
	return nil
}

// ServerXMLHTTP wraps MsXML2ServerXMLHTTP for ASPLibrary interface compatibility
type ServerXMLHTTP struct {
	lib *MsXML2ServerXMLHTTP
}

// NewServerXMLHTTP creates a new ServerXMLHTTP instance
func NewServerXMLHTTP(ctx *ExecutionContext) *ServerXMLHTTP {
	return &ServerXMLHTTP{
		lib: NewMsXML2ServerXMLHTTP(ctx),
	}
}

// CallMethod calls a method on ServerXMLHTTP
func (sx *ServerXMLHTTP) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return sx.lib.CallMethod(name, args), nil
}

// GetProperty gets a property from ServerXMLHTTP
func (sx *ServerXMLHTTP) GetProperty(name string) interface{} {
	return sx.lib.GetProperty(name)
}

// SetProperty sets a property on ServerXMLHTTP
func (sx *ServerXMLHTTP) SetProperty(name string, value interface{}) error {
	sx.lib.SetProperty(name, value)
	return nil
}

// DOMDocument wraps MsXML2DOMDocument for ASPLibrary interface compatibility
type DOMDocument struct {
	lib *MsXML2DOMDocument
}

// NewDOMDocument creates a new DOMDocument instance
func NewDOMDocument(ctx *ExecutionContext) *DOMDocument {
	return &DOMDocument{
		lib: NewMsXML2DOMDocument(ctx),
	}
}

// CallMethod calls a method on DOMDocument
func (dd *DOMDocument) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return dd.lib.CallMethod(name, args), nil
}

// GetProperty gets a property from DOMDocument
func (dd *DOMDocument) GetProperty(name string) interface{} {
	return dd.lib.GetProperty(name)
}

// SetProperty sets a property on DOMDocument
func (dd *DOMDocument) SetProperty(name string, value interface{}) error {
	dd.lib.SetProperty(name, value)
	return nil
}

// ADOConnection wraps ADODBConnection for ASPLibrary interface compatibility
type ADOConnection struct {
	lib *ADODBConnection
}

// NewADOConnection creates a new ADOConnection instance
func NewADOConnection(ctx *ExecutionContext) *ADOConnection {
	return &ADOConnection{
		lib: NewADODBConnection(ctx),
	}
}

// CallMethod calls a method on ADOConnection
func (ac *ADOConnection) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return ac.lib.CallMethod(name, args), nil
}

// GetProperty gets a property from ADOConnection
func (ac *ADOConnection) GetProperty(name string) interface{} {
	return ac.lib.GetProperty(name)
}

// SetProperty sets a property on ADOConnection
func (ac *ADOConnection) SetProperty(name string, value interface{}) error {
	ac.lib.SetProperty(name, value)
	return nil
}

// ADORecordset wraps ADODBRecordset for ASPLibrary interface compatibility
type ADORecordset struct {
	lib *ADODBRecordset
}

// NewADORecordset creates a new ADORecordset instance
func NewADORecordset(ctx *ExecutionContext) *ADORecordset {
	return &ADORecordset{
		lib: NewADODBRecordset(ctx),
	}
}

// CallMethod calls a method on ADORecordset
func (ar *ADORecordset) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return ar.lib.CallMethod(name, args), nil
}

// GetProperty gets a property from ADORecordset
func (ar *ADORecordset) GetProperty(name string) interface{} {
	return ar.lib.GetProperty(name)
}

// SetProperty sets a property on ADORecordset
func (ar *ADORecordset) SetProperty(name string, value interface{}) error {
	ar.lib.SetProperty(name, value)
	return nil
}

// ADOStream wraps ADODBStream for ASPLibrary interface compatibility
type ADOStream struct {
	lib *ADODBStream
}

// NewADOStream creates a new ADOStream instance
func NewADOStream(ctx *ExecutionContext) *ADOStream {
	return &ADOStream{
		lib: NewADODBStream(ctx),
	}
}

// CallMethod calls a method on ADOStream
func (as *ADOStream) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return as.lib.CallMethod(name, args), nil
}

// GetProperty gets a property from ADOStream
func (as *ADOStream) GetProperty(name string) interface{} {
	return as.lib.GetProperty(name)
}

// SetProperty sets a property on ADOStream
func (as *ADOStream) SetProperty(name string, value interface{}) error {
	as.lib.SetProperty(name, value)
	return nil
}
