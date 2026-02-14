/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
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
	return jl.lib.CallMethod(name, args...), nil
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
	return fs.lib.CallMethod(name, args...), nil
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

// FileSystemObjectLibrary wraps FSOObject (Scripting.FileSystemObject) for ASPLibrary interface compatibility
type FileSystemObjectLibrary struct {
	fso *FSOObject
}

// NewFileSystemObjectLibrary creates a new FileSystemObject library instance (Scripting.FileSystemObject)
func NewFileSystemObjectLibrary(ctx *ExecutionContext) *FileSystemObjectLibrary {
	return &FileSystemObjectLibrary{
		fso: &FSOObject{ctx: ctx},
	}
}

// CallMethod calls a method on the FileSystemObject library
func (fol *FileSystemObjectLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return fol.fso.CallMethod(name, args...), nil
}

// GetProperty gets a property from the FileSystemObject library
func (fol *FileSystemObjectLibrary) GetProperty(name string) interface{} {
	return fol.fso.GetProperty(name)
}

// SetProperty sets a property on the FileSystemObject library
func (fol *FileSystemObjectLibrary) SetProperty(name string, value interface{}) error {
	fol.fso.SetProperty(name, value)
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
	return hl.lib.CallMethod(name, args...), nil
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
	return tl.lib.CallMethod(name, args...), nil
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
	return ml.lib.CallMethod(name, args...), nil
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
	return cl.lib.CallMethod(name, args...), nil
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

// GetName returns the name of the object
func (sx *ServerXMLHTTP) GetName() string {
	return "MSXML2.ServerXMLHTTP"
}

// CallMethod calls a method on ServerXMLHTTP
func (sx *ServerXMLHTTP) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return sx.lib.CallMethod(name, args...)
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

func (dd *DOMDocument) GetName() string {
	return "MSXML2.DOMDocument"
}

// NewDOMDocument creates a new DOMDocument instance
func NewDOMDocument(ctx *ExecutionContext) *DOMDocument {
	return &DOMDocument{
		lib: NewMsXML2DOMDocument(ctx),
	}
}

// CallMethod calls a method on DOMDocument
func (dd *DOMDocument) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return dd.lib.CallMethod(name, args...)
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
	return ac.lib.CallMethod(name, args...), nil
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

// GetName returns the name of the object
func (ac *ADOConnection) GetName() string {
	return "ADODB.Connection"
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
	return ar.lib.CallMethod(name, args...), nil
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

// GetName returns the name of the object
func (ar *ADORecordset) GetName() string {
	return "ADODB.Recordset"
}

// ADOOLERecordset wraps ADODBOLERecordset for ASPLibrary interface compatibility
type ADOOLERecordset struct {
	lib *ADODBOLERecordset
}

// NewADOOLERecordset creates a new ADOOLERecordset instance
func NewADOOLERecordset(oleRs *ADODBOLERecordset) *ADOOLERecordset {
	return &ADOOLERecordset{
		lib: oleRs,
	}
}

// CallMethod calls a method on ADOOLERecordset
func (ar *ADOOLERecordset) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return ar.lib.CallMethod(name, args...), nil
}

// GetProperty gets a property from ADOOLERecordset
func (ar *ADOOLERecordset) GetProperty(name string) interface{} {
	return ar.lib.GetProperty(name)
}

// SetProperty sets a property on ADOOLERecordset
func (ar *ADOOLERecordset) SetProperty(name string, value interface{}) error {
	ar.lib.SetProperty(name, value)
	return nil
}

// GetName returns the name of the object
func (ar *ADOOLERecordset) GetName() string {
	return "ADODB.Recordset"
}

// ADOOLEFields wraps ADODBOLEFields for ASPLibrary interface compatibility
type ADOOLEFields struct {
	lib *ADODBOLEFields
}

// NewADOOLEFields creates a new ADOOLEFields instance
func NewADOOLEFields(oleFields *ADODBOLEFields) *ADOOLEFields {
	return &ADOOLEFields{
		lib: oleFields,
	}
}

// CallMethod calls a method on ADOOLEFields
func (af *ADOOLEFields) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return af.lib.CallMethod(name, args...), nil
}

// GetProperty gets a property from ADOOLEFields
func (af *ADOOLEFields) GetProperty(name string) interface{} {
	return af.lib.GetProperty(name)
}

// SetProperty sets a property on ADOOLEFields
func (af *ADOOLEFields) SetProperty(name string, value interface{}) error {
	af.lib.SetProperty(name, value)
	return nil
}

// Enumeration returns all Field proxies for For Each support.
func (af *ADOOLEFields) Enumeration() []interface{} {
	if af.lib == nil {
		return []interface{}{}
	}
	return af.lib.Enumeration()
}

// GetName returns the name of the object
func (af *ADOOLEFields) GetName() string {
	return "ADODB.Fields"
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
	return as.lib.CallMethod(name, args...), nil
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

// GetName returns the name of the object
func (as *ADOStream) GetName() string {
	return "ADODB.Stream"
}

// ADOCommand wraps ADODBCommand for ASPLibrary interface compatibility
type ADOCommand struct {
	lib *ADODBCommand
}

// NewADOCommand creates a new ADOCommand instance
func NewADOCommand(ctx *ExecutionContext) *ADOCommand {
	return &ADOCommand{
		lib: NewADODBCommand(ctx),
	}
}

// CallMethod calls a method on ADOCommand
func (ac *ADOCommand) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return ac.lib.CallMethod(name, args...), nil
}

// GetProperty gets a property from ADOCommand
func (ac *ADOCommand) GetProperty(name string) interface{} {
	return ac.lib.GetProperty(name)
}

// SetProperty sets a property on ADOCommand
func (ac *ADOCommand) SetProperty(name string, value interface{}) error {
	ac.lib.SetProperty(name, value)
	return nil
}

// GetName returns the name of the object
func (ac *ADOCommand) GetName() string {
	return "ADODB.Command"
}

// RegExpLibrary wraps G3REGEXP for ASPLibrary interface compatibility
type RegExpLibrary struct {
	lib *G3REGEXP
}

// NewRegExpLibrary creates a new RegExp library instance
func NewRegExpLibrary(ctx *ExecutionContext) *RegExpLibrary {
	return &RegExpLibrary{
		lib: &G3REGEXP{},
	}
}

// CallMethod calls a method on the RegExp library
func (rl *RegExpLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return rl.lib.CallMethod(name, args...)
}

// GetProperty gets a property from the RegExp library
func (rl *RegExpLibrary) GetProperty(name string) interface{} {
	return rl.lib.GetProperty(name)
}

// SetProperty sets a property on the RegExp library
func (rl *RegExpLibrary) SetProperty(name string, value interface{}) error {
	return rl.lib.SetProperty(name, value)
}

type WScriptShellLibrary struct {
	lib *WScriptShell
}

// NewWScriptShellLibrary creates a new WScriptShell library instance
func NewWScriptShellLibrary(ctx *ExecutionContext) *WScriptShellLibrary {
	return &WScriptShellLibrary{
		lib: NewWScriptShell(ctx),
	}
}

// CallMethod calls a method on the WScriptShell library
func (wsl *WScriptShellLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return wsl.lib.CallMethod(name, args...), nil
}

// GetProperty gets a property from the WScriptShell library
func (wsl *WScriptShellLibrary) GetProperty(name string) interface{} {
	return wsl.lib.GetProperty(name)
}

// SetProperty sets a property on the WScriptShell library
func (wsl *WScriptShellLibrary) SetProperty(name string, value interface{}) error {
	wsl.lib.SetProperty(name, value)
	return nil
}

// G3DBLibrary wraps G3DB for ASPLibrary interface compatibility
type G3DBLibrary struct {
	lib *G3DB
}

// NewG3DBLibrary creates a new G3DB library instance
func NewG3DBLibrary(ctx *ExecutionContext) *G3DBLibrary {
	return &G3DBLibrary{
		lib: NewG3DB(ctx),
	}
}

// CallMethod calls a method on the G3DB library
func (dbl *G3DBLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return dbl.lib.CallMethod(name, args...), nil
}

// GetProperty gets a property from the G3DB library
func (dbl *G3DBLibrary) GetProperty(name string) interface{} {
	return dbl.lib.GetProperty(name)
}

// SetProperty sets a property on the G3DB library
func (dbl *G3DBLibrary) SetProperty(name string, value interface{}) error {
	dbl.lib.SetProperty(name, value)
	return nil
}

// ZIPLibrary wraps G3ZIP for ASPLibrary interface compatibility
type ZIPLibrary struct {
	lib *G3ZIP
}

// NewZIPLibrary creates a new ZIP library instance
func NewZIPLibrary(ctx *ExecutionContext) *ZIPLibrary {
	return &ZIPLibrary{
		lib: &G3ZIP{ctx: ctx},
	}
}

// CallMethod calls a method on the ZIP library
func (zl *ZIPLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return zl.lib.CallMethod(name, args...), nil
}

// GetProperty gets a property from the ZIP library
func (zl *ZIPLibrary) GetProperty(name string) interface{} {
	return zl.lib.GetProperty(name)
}

// SetProperty sets a property on the ZIP library
func (zl *ZIPLibrary) SetProperty(name string, value interface{}) error {
	zl.lib.SetProperty(name, value)
	return nil
}

// G3FCLibrary wraps G3FC for ASPLibrary interface compatibility
type G3FCLibrary struct {
	lib *G3FC
}

// ImageLibrary wraps G3IMAGE for ASPLibrary interface compatibility
type ImageLibrary struct {
	lib *G3IMAGE
}

// PDFLibrary wraps G3PDF for ASPLibrary interface compatibility
type PDFLibrary struct {
	lib *G3PDF
}

// NewG3FCLibrary creates a new G3FC library instance
func NewG3FCLibrary(ctx *ExecutionContext) *G3FCLibrary {
	return &G3FCLibrary{
		lib: &G3FC{ctx: ctx},
	}
}

// NewImageLibrary creates a new Image library instance
func NewImageLibrary(ctx *ExecutionContext) *ImageLibrary {
	return &ImageLibrary{
		lib: &G3IMAGE{ctx: ctx},
	}
}

// NewPDFLibrary creates a new PDF library instance
func NewPDFLibrary(ctx *ExecutionContext) *PDFLibrary {
	return &PDFLibrary{
		lib: NewG3PDF(ctx),
	}
}

// CallMethod calls a method on the G3FC library
func (gl *G3FCLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return gl.lib.CallMethod(name, args...), nil
}

// GetProperty gets a property from the G3FC library
func (gl *G3FCLibrary) GetProperty(name string) interface{} {
	return gl.lib.GetProperty(name)
}

// SetProperty sets a property on the G3FC library
func (gl *G3FCLibrary) SetProperty(name string, value interface{}) error {
	gl.lib.SetProperty(name, value)
	return nil
}

// CallMethod calls a method on the Image library
func (il *ImageLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return il.lib.CallMethod(name, args...), nil
}

// GetProperty gets a property from the Image library
func (il *ImageLibrary) GetProperty(name string) interface{} {
	return il.lib.GetProperty(name)
}

// SetProperty sets a property on the Image library
func (il *ImageLibrary) SetProperty(name string, value interface{}) error {
	il.lib.SetProperty(name, value)
	return nil
}

// CallMethod calls a method on the PDF library
func (pl *PDFLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return pl.lib.CallMethod(name, args...), nil
}

// GetProperty gets a property from the PDF library
func (pl *PDFLibrary) GetProperty(name string) interface{} {
	return pl.lib.GetProperty(name)
}

// SetProperty sets a property on the PDF library
func (pl *PDFLibrary) SetProperty(name string, value interface{}) error {
	pl.lib.SetProperty(name, value)
	return nil
}
