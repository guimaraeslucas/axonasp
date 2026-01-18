package server

import (
	"fmt"
	"net/http"
	"strings"
	"sync"
	"time"
)

// ResponseObject represents the ASP Classic Response Object
// Implements all methods, properties, and collections from Classic ASP
type ResponseObject struct {
	// Internal state
	buffer          []byte
	httpWriter      http.ResponseWriter
	httpRequest     *http.Request
	isEnded         bool
	isFlushed       bool
	headers         map[string]string
	cookiesMap      map[string]*ResponseCookie
	logEntries      []string
	mu              sync.RWMutex

	// Properties
	bufferEnabled   bool
	cacheControl    string
	charset         string
	contentType     string
	expires         int       // Minutes from now
	expiresAbsolute time.Time // Absolute expiration time
	pics            string    // PICS label
	status          string    // HTTP status (e.g., "200 OK")
}

// ResponseCookie represents a cookie in the Response.Cookies collection
type ResponseCookie struct {
	Name     string
	Value    string
	Domain   string
	Path     string
	Expires  time.Time
	Secure   bool
	HttpOnly bool
}

// ResponseCookiesCollection represents the Response.Cookies collection
type ResponseCookiesCollection struct {
	response *ResponseObject
}

// NewResponseObject creates a new Response object with default values
func NewResponseObject(w http.ResponseWriter, r *http.Request) *ResponseObject {
	return &ResponseObject{
		buffer:        make([]byte, 0),
		httpWriter:    w,
		httpRequest:   r,
		headers:       make(map[string]string),
		cookiesMap:    make(map[string]*ResponseCookie),
		logEntries:    make([]string, 0),
		bufferEnabled: true,
		cacheControl:  "Private",
		charset:       "utf-8",
		contentType:   "text/html",
		status:        "200 OK",
	}
}

// ==================== METHODS ====================

// Write outputs content to the HTTP response
// Usage: Response.Write(data)
func (r *ResponseObject) Write(data interface{}) error {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.isEnded {
		return nil
	}

	str := r.toString(data)
	r.buffer = append(r.buffer, []byte(str)...)

	// If buffering is disabled, flush immediately
	if !r.bufferEnabled {
		return r.flushInternal()
	}

	return nil
}

// BinaryWrite outputs binary data to the HTTP response
// Usage: Response.BinaryWrite(data)
func (r *ResponseObject) BinaryWrite(data interface{}) error {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.isEnded {
		return nil
	}

	var bytes []byte
	switch v := data.(type) {
	case []byte:
		bytes = v
	case string:
		bytes = []byte(v)
	default:
		bytes = []byte(fmt.Sprintf("%v", v))
	}

	r.buffer = append(r.buffer, bytes...)

	// If buffering is disabled, flush immediately
	if !r.bufferEnabled {
		return r.flushInternal()
	}

	return nil
}

// AddHeader adds an HTTP header to the response
// Usage: Response.AddHeader(name, value)
func (r *ResponseObject) AddHeader(name, value string) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.isFlushed || r.isEnded {
		return
	}

	r.headers[name] = value
}

// AppendToLog adds a string to the web server log
// Usage: Response.AppendToLog(message)
func (r *ResponseObject) AppendToLog(message string) {
	r.mu.Lock()
	defer r.mu.Unlock()

	r.logEntries = append(r.logEntries, message)
	// In production, this would write to the actual web server log
	fmt.Printf("[ASP Log] %s\n", message)
}

// Clear clears the response buffer
// Usage: Response.Clear()
func (r *ResponseObject) Clear() {
	r.mu.Lock()
	defer r.mu.Unlock()

	if !r.isEnded {
		r.buffer = make([]byte, 0)
	}
}

// Flush sends buffered output immediately
// Usage: Response.Flush()
func (r *ResponseObject) Flush() error {
	r.mu.Lock()
	defer r.mu.Unlock()

	return r.flushInternal()
}

// End stops processing the ASP page and sends the output
// Usage: Response.End()
func (r *ResponseObject) End() error {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.isEnded {
		return nil
	}

	err := r.flushInternal()
	r.isEnded = true
	return err
}

// Redirect redirects the client to a different URL
// Usage: Response.Redirect(url)
func (r *ResponseObject) Redirect(url string) error {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.isFlushed || r.isEnded {
		return nil
	}

	// Clear buffer before redirect
	r.buffer = make([]byte, 0)

	// Set redirect header
	r.httpWriter.Header().Set("Location", url)
	r.httpWriter.WriteHeader(http.StatusFound)
	r.isFlushed = true
	r.isEnded = true

	return nil
}

// ==================== PROPERTIES ====================

// GetBuffer gets the Buffer property (enable/disable response buffering)
func (r *ResponseObject) GetBuffer() bool {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.bufferEnabled
}

// SetBuffer sets the Buffer property
func (r *ResponseObject) SetBuffer(enabled bool) {
	r.mu.Lock()
	defer r.mu.Unlock()
	r.bufferEnabled = enabled
}

// GetCacheControl gets the Cache-Control header value
func (r *ResponseObject) GetCacheControl() string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.cacheControl
}

// SetCacheControl sets the Cache-Control header
func (r *ResponseObject) SetCacheControl(value string) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if !r.isFlushed {
		r.cacheControl = value
	}
}

// GetCharset gets the charset for the Content-Type header
func (r *ResponseObject) GetCharset() string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.charset
}

// SetCharset sets the charset for the Content-Type header
func (r *ResponseObject) SetCharset(value string) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if !r.isFlushed {
		r.charset = value
	}
}

// GetContentType gets the Content-Type header value
func (r *ResponseObject) GetContentType() string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.contentType
}

// SetContentType sets the Content-Type header
func (r *ResponseObject) SetContentType(value string) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if !r.isFlushed {
		r.contentType = value
	}
}

// GetExpires gets the Expires property (minutes from now)
func (r *ResponseObject) GetExpires() int {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.expires
}

// SetExpires sets the Expires property (minutes from now)
func (r *ResponseObject) SetExpires(minutes int) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if !r.isFlushed {
		r.expires = minutes
	}
}

// GetExpiresAbsolute gets the ExpiresAbsolute property
func (r *ResponseObject) GetExpiresAbsolute() time.Time {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.expiresAbsolute
}

// SetExpiresAbsolute sets the ExpiresAbsolute property
func (r *ResponseObject) SetExpiresAbsolute(t time.Time) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if !r.isFlushed {
		r.expiresAbsolute = t
	}
}

// IsClientConnected checks if the client is still connected
func (r *ResponseObject) IsClientConnected() bool {
	if r.httpRequest == nil {
		return true
	}

	select {
	case <-r.httpRequest.Context().Done():
		return false
	default:
		return true
	}
}

// GetPICS gets the PICS label
func (r *ResponseObject) GetPICS() string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.pics
}

// SetPICS sets the PICS label
func (r *ResponseObject) SetPICS(value string) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if !r.isFlushed {
		r.pics = value
	}
}

// GetStatus gets the HTTP status line
func (r *ResponseObject) GetStatus() string {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.status
}

// SetStatus sets the HTTP status line (e.g., "404 Not Found")
func (r *ResponseObject) SetStatus(value string) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if !r.isFlushed {
		r.status = value
	}
}

// ==================== COLLECTIONS ====================

// Cookies returns the Response.Cookies collection
func (r *ResponseObject) Cookies() *ResponseCookiesCollection {
	return &ResponseCookiesCollection{response: r}
}

// GetCookie gets a cookie by name (for internal use)
func (r *ResponseObject) GetCookie(name string) *ResponseCookie {
	r.mu.RLock()
	defer r.mu.RUnlock()

	if cookie, exists := r.cookiesMap[strings.ToLower(name)]; exists {
		return cookie
	}
	return nil
}

// SetCookie sets a cookie value
func (r *ResponseObject) SetCookie(name, value string) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.isFlushed {
		return
	}

	key := strings.ToLower(name)
	if cookie, exists := r.cookiesMap[key]; exists {
		cookie.Value = value
	} else {
		r.cookiesMap[key] = &ResponseCookie{
			Name:  name,
			Value: value,
			Path:  "/",
		}
	}
}

// SetCookieWithOptions sets a cookie with all options
func (r *ResponseObject) SetCookieWithOptions(cookie *ResponseCookie) {
	r.mu.Lock()
	defer r.mu.Unlock()

	if r.isFlushed {
		return
	}

	key := strings.ToLower(cookie.Name)
	r.cookiesMap[key] = cookie
}

// ==================== INTERNAL METHODS ====================

// flushInternal sends the buffered output (must be called with lock held)
func (r *ResponseObject) flushInternal() error {
	if r.isFlushed || r.httpWriter == nil {
		return nil
	}

	// Set Content-Type with charset
	ct := r.contentType
	if r.charset != "" && !strings.Contains(strings.ToLower(ct), "charset") {
		ct += "; charset=" + r.charset
	}
	if ct != "" {
		r.httpWriter.Header().Set("Content-Type", ct)
	}

	// Set Cache-Control
	if r.cacheControl != "" {
		r.httpWriter.Header().Set("Cache-Control", r.cacheControl)

		// Set Pragma for no-cache scenarios
		if strings.EqualFold(r.cacheControl, "no-cache") || 
		   (strings.EqualFold(r.cacheControl, "private") && r.expires <= 0) {
			r.httpWriter.Header().Set("Pragma", "no-cache")
		}
	}

	// Set Expires header
	if r.expires > 0 {
		expTime := time.Now().Add(time.Duration(r.expires) * time.Minute)
		r.httpWriter.Header().Set("Expires", expTime.Format(http.TimeFormat))
	} else if r.expires == 0 {
		r.httpWriter.Header().Set("Expires", "0")
	} else if r.expires < 0 {
		// Negative values mean immediate expiration
		r.httpWriter.Header().Set("Expires", time.Now().Add(-1*time.Hour).Format(http.TimeFormat))
	}

	// Set ExpiresAbsolute
	if !r.expiresAbsolute.IsZero() {
		r.httpWriter.Header().Set("Expires", r.expiresAbsolute.Format(http.TimeFormat))
	}

	// Set PICS label
	if r.pics != "" {
		r.httpWriter.Header().Set("PICS-Label", r.pics)
	}

	// Set custom headers
	for name, value := range r.headers {
		r.httpWriter.Header().Set(name, value)
	}

	// Set cookies
	for _, cookie := range r.cookiesMap {
		httpCookie := &http.Cookie{
			Name:     cookie.Name,
			Value:    cookie.Value,
			Path:     cookie.Path,
			Domain:   cookie.Domain,
			Expires:  cookie.Expires,
			Secure:   cookie.Secure,
			HttpOnly: cookie.HttpOnly,
		}
		if cookie.Path == "" {
			httpCookie.Path = "/"
		}
		http.SetCookie(r.httpWriter, httpCookie)
	}

	// Set status code
	if r.status != "" && r.status != "200 OK" {
		// Parse status code from status string (e.g., "404 Not Found")
		parts := strings.SplitN(r.status, " ", 2)
		if len(parts) > 0 {
			if code, err := fmt.Sscanf(parts[0], "%d", new(int)); err == nil && code == 1 {
				var statusCode int
				fmt.Sscanf(parts[0], "%d", &statusCode)
				r.httpWriter.WriteHeader(statusCode)
			}
		}
	}

	// Write buffer
	if len(r.buffer) > 0 {
		_, err := r.httpWriter.Write(r.buffer)
		if err != nil {
			return err
		}
	}

	r.isFlushed = true
	return nil
}

// toString converts a value to string following ASP rules
func (r *ResponseObject) toString(value interface{}) string {
	if value == nil {
		return ""
	}

	switch v := value.(type) {
	case string:
		return v
	case int, int32, int64:
		return fmt.Sprintf("%d", v)
	case float32, float64:
		return fmt.Sprintf("%v", v)
	case bool:
		if v {
			return "True"
		}
		return "False"
	case []byte:
		return string(v)
	default:
		return fmt.Sprintf("%v", v)
	}
}

// GetBufferContent returns the current buffer content (for testing/debugging)
func (r *ResponseObject) GetBufferContent() []byte {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.buffer
}

// IsEnded checks if Response.End() has been called
func (r *ResponseObject) IsEnded() bool {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return r.isEnded
}

// ==================== COOKIES COLLECTION METHODS ====================

// Item gets or sets a cookie value
// Usage: Response.Cookies("name") or Response.Cookies("name") = "value"
func (c *ResponseCookiesCollection) Item(name string) string {
	cookie := c.response.GetCookie(name)
	if cookie != nil {
		return cookie.Value
	}
	return ""
}

// SetItem sets a cookie value
func (c *ResponseCookiesCollection) SetItem(name, value string) {
	c.response.SetCookie(name, value)
}

// SetDomain sets the domain for a cookie
// Usage: Response.Cookies("name").Domain = "example.com"
func (c *ResponseCookiesCollection) SetDomain(name, domain string) {
	c.response.mu.Lock()
	defer c.response.mu.Unlock()

	key := strings.ToLower(name)
	if cookie, exists := c.response.cookiesMap[key]; exists {
		cookie.Domain = domain
	} else {
		c.response.cookiesMap[key] = &ResponseCookie{
			Name:   name,
			Domain: domain,
			Path:   "/",
		}
	}
}

// SetPath sets the path for a cookie
// Usage: Response.Cookies("name").Path = "/app"
func (c *ResponseCookiesCollection) SetPath(name, path string) {
	c.response.mu.Lock()
	defer c.response.mu.Unlock()

	key := strings.ToLower(name)
	if cookie, exists := c.response.cookiesMap[key]; exists {
		cookie.Path = path
	} else {
		c.response.cookiesMap[key] = &ResponseCookie{
			Name: name,
			Path: path,
		}
	}
}

// SetExpires sets the expiration time for a cookie
// Usage: Response.Cookies("name").Expires = #2025/12/31#
func (c *ResponseCookiesCollection) SetExpires(name string, expires time.Time) {
	c.response.mu.Lock()
	defer c.response.mu.Unlock()

	key := strings.ToLower(name)
	if cookie, exists := c.response.cookiesMap[key]; exists {
		cookie.Expires = expires
	} else {
		c.response.cookiesMap[key] = &ResponseCookie{
			Name:    name,
			Expires: expires,
			Path:    "/",
		}
	}
}

// SetSecure sets the Secure flag for a cookie
// Usage: Response.Cookies("name").Secure = True
func (c *ResponseCookiesCollection) SetSecure(name string, secure bool) {
	c.response.mu.Lock()
	defer c.response.mu.Unlock()

	key := strings.ToLower(name)
	if cookie, exists := c.response.cookiesMap[key]; exists {
		cookie.Secure = secure
	} else {
		c.response.cookiesMap[key] = &ResponseCookie{
			Name:   name,
			Secure: secure,
			Path:   "/",
		}
	}
}

// SetHttpOnly sets the HttpOnly flag for a cookie
// Usage: Response.Cookies("name").HttpOnly = True
func (c *ResponseCookiesCollection) SetHttpOnly(name string, httpOnly bool) {
	c.response.mu.Lock()
	defer c.response.mu.Unlock()

	key := strings.ToLower(name)
	if cookie, exists := c.response.cookiesMap[key]; exists {
		cookie.HttpOnly = httpOnly
	} else {
		c.response.cookiesMap[key] = &ResponseCookie{
			Name:     name,
			HttpOnly: httpOnly,
			Path:     "/",
		}
	}
}
