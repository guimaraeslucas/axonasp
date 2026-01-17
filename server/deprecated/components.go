package asp

import (
	"fmt"
	"io"
	"net/http"
	"regexp"
	"strings"
	"sync"
)

// Component is the interface that all COM-like objects must implement
// to interact with the interpreter's dynamic evaluation.
type Component interface {
	GetProperty(name string) interface{}
	SetProperty(name string, value interface{})
	CallMethod(name string, args []interface{}) interface{}
}

// Enumerable interface allows For Each support in custom components
type Enumerable interface {
	Enumeration() []interface{}
}

// ComponentFactory creates a new component based on ProgID
func ComponentFactory(progID string, ctx *ExecutionContext) interface{} {
	// Debug logging
	//fmt.Printf("ComponentFactory requested: %s\n", progID)

	switch strings.ToLower(progID) {
	case "scripting.dictionary":
		return NewDictionary()
	case "msxml2.xmlhttp", "microsoft.xmlhttp":
		return NewXMLHTTP()
	case "msxml2.serverxmlhttp":
		return NewMsXML2ServerXMLHTTP(ctx)
	case "msxml2.domdocument":
		return NewMsXML2DOMDocument(ctx)
	case "adodb.connection":
		return NewADODBConnection(ctx)
	case "adodb.recordset":
		return NewADODBRecordset(ctx)
	case "adodb.stream":
		return NewADODBStream(ctx)
	case "g3json":
		return &G3JSON{}
	case "g3files":
		return &G3FILES{ctx: ctx}
	case "g3http":
		return &G3HTTP{ctx: ctx}
	case "g3template":
		return &G3TEMPLATE{ctx: ctx}
	case "g3mail":
		return &G3MAIL{ctx: ctx}
	case "g3crypto":
		return &G3CRYPTO{ctx: ctx}
	case "vbscript.regexp":
		return NewRegExpObject()
	case "scripting.filesystemobject":
		return &FSOObject{ctx: ctx}
	default:
		return nil
	}
}

// --- Scripting.Dictionary ---

type Dictionary struct {
	store map[string]interface{}
	mutex sync.RWMutex
}

func NewDictionary() *Dictionary {
	return &Dictionary{
		store: make(map[string]interface{}),
	}
}

func (d *Dictionary) GetProperty(name string) interface{} {
	d.mutex.RLock()
	defer d.mutex.RUnlock()

	if strings.EqualFold(name, "Count") {
		return len(d.store)
	}
	return nil
}

func (d *Dictionary) SetProperty(name string, value interface{}) {
	// Properties like CompareMode are often ignored in simple impl
}

func (d *Dictionary) CallMethod(name string, args []interface{}) interface{} {
	d.mutex.Lock()
	defer d.mutex.Unlock()

	switch strings.ToLower(name) {
	case "add":
		if len(args) >= 2 {
			key := fmt.Sprintf("%v", args[0])
			d.store[key] = args[1]
		}
	case "remove":
		if len(args) >= 1 {
			key := fmt.Sprintf("%v", args[0])
			delete(d.store, key)
		}
	case "removeall":
		d.store = make(map[string]interface{})
	case "exists":
		if len(args) >= 1 {
			key := fmt.Sprintf("%v", args[0])
			_, exists := d.store[key]
			return exists
		}
	case "item":
		if len(args) >= 1 {
			key := fmt.Sprintf("%v", args[0])
			return d.store[key]
		}
	case "keys":
		keys := make([]interface{}, 0, len(d.store))
		for k := range d.store {
			keys = append(keys, k)
		}
		return keys
	case "items":
		items := make([]interface{}, 0, len(d.store))
		for _, v := range d.store {
			items = append(items, v)
		}
		return items
	}
	return nil
}

func (d *Dictionary) Enumeration() []interface{} {
	d.mutex.RLock()
	defer d.mutex.RUnlock()
	keys := make([]interface{}, 0, len(d.store))
	for k := range d.store {
		keys = append(keys, k)
	}
	return keys
}

// --- MSXML2.XMLHTTP ---

type XMLHTTP struct {
	method       string
	url          string
	responseText string
	status       int
	headers      map[string]string
}

func NewXMLHTTP() *XMLHTTP {
	return &XMLHTTP{
		headers: make(map[string]string),
	}
}

func (x *XMLHTTP) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "responsetext":
		return x.responseText
	case "status":
		return x.status
	case "readystate":
		return 4 // Always complete in this synchronous shim
	}
	return nil
}

func (x *XMLHTTP) SetProperty(name string, value interface{}) {
	// Read-only mostly
}

func (x *XMLHTTP) CallMethod(name string, args []interface{}) interface{} {
	switch strings.ToLower(name) {
	case "open":
		// Open(Method, URL, [Async], [User], [Password])
		if len(args) >= 2 {
			x.method = fmt.Sprintf("%v", args[0])
			x.url = fmt.Sprintf("%v", args[1])
		}
	case "setrequestheader":
		if len(args) >= 2 {
			k := fmt.Sprintf("%v", args[0])
			v := fmt.Sprintf("%v", args[1])
			x.headers[k] = v
		}
	case "send":
		var body io.Reader
		if len(args) > 0 && args[0] != nil {
			bodyStr := fmt.Sprintf("%v", args[0])
			body = strings.NewReader(bodyStr)
		}

		req, err := http.NewRequest(x.method, x.url, body)
		if err != nil {
			x.status = 500
			x.responseText = err.Error()
			return nil
		}

		for k, v := range x.headers {
			req.Header.Set(k, v)
		}

		client := &http.Client{}
		resp, err := client.Do(req)
		if err != nil {
			x.status = 500
			x.responseText = err.Error()
			return nil
		}
		defer resp.Body.Close()

		bodyBytes, _ := io.ReadAll(resp.Body)
		x.status = resp.StatusCode
		x.responseText = string(bodyBytes)
	}
	return nil
}

// ADODB implementations moved to database_lib.go

// --- VBScript.RegExp ---

type RegExpObject struct {
	Pattern    string
	IgnoreCase bool
	Global     bool
	MultiLine  bool
}

func NewRegExpObject() *RegExpObject {
	return &RegExpObject{
		Pattern:    "",
		IgnoreCase: false,
		Global:     false,
		MultiLine:  false,
	}
}

func (r *RegExpObject) compile() (*regexp.Regexp, error) {
	pattern := r.Pattern
	flags := ""
	if r.IgnoreCase {
		flags += "i"
	}
	if r.MultiLine {
		flags += "m"
	}
	if flags != "" {
		pattern = "(?" + flags + ")" + pattern
	}
	return regexp.Compile(pattern)
}

func (r *RegExpObject) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "pattern":
		return r.Pattern
	case "ignorecase":
		return r.IgnoreCase
	case "global":
		return r.Global
	case "multiline":
		return r.MultiLine
	}
	return nil
}

func (r *RegExpObject) SetProperty(name string, value interface{}) {
	switch strings.ToLower(name) {
	case "pattern":
		r.Pattern = fmt.Sprintf("%v", value)
	case "ignorecase":
		r.IgnoreCase = fmt.Sprintf("%v", value) == "true" || value == true
	case "global":
		r.Global = fmt.Sprintf("%v", value) == "true" || value == true
	case "multiline":
		r.MultiLine = fmt.Sprintf("%v", value) == "true" || value == true
	}
}

func (r *RegExpObject) CallMethod(name string, args []interface{}) interface{} {
	re, err := r.compile()
	if err != nil {
		// In VBScript this would be a runtime error
		return nil
	}

	switch strings.ToLower(name) {
	case "execute":
		if len(args) > 0 {
			input := fmt.Sprintf("%v", args[0])
			var matches []interface{}
			if r.Global {
				// Find all
				found := re.FindAllStringIndex(input, -1)
				for _, loc := range found {
					matchVal := input[loc[0]:loc[1]]
					matches = append(matches, &MatchObject{
						Value:      matchVal,
						FirstIndex: loc[0],
						Length:     len(matchVal),
					})
				}
			} else {
				// Find first
				loc := re.FindStringIndex(input)
				if loc != nil {
					matchVal := input[loc[0]:loc[1]]
					matches = append(matches, &MatchObject{
						Value:      matchVal,
						FirstIndex: loc[0],
						Length:     len(matchVal),
					})
				}
			}
			return &MatchesCollection{Matches: matches}
		}
	case "test":
		if len(args) > 0 {
			input := fmt.Sprintf("%v", args[0])
			return re.MatchString(input)
		}
	case "replace":
		if len(args) >= 2 {
			input := fmt.Sprintf("%v", args[0])
			repl := fmt.Sprintf("%v", args[1])
			// VBScript replace logic:
			// If Global = True, ReplaceAll
			// If Global = False, Replace First (Go doesn't support ReplaceFirst natively in regexp easily without Split)
			// Wait, regexp.ReplaceAllString replaces all.
			// To replace only first, we might need manual logic if Global is false.
			if r.Global {
				return re.ReplaceAllString(input, repl)
			} else {
				// Replace only the first occurrence
				loc := re.FindStringIndex(input)
				if loc == nil {
					return input
				}
				// Construct result: before + repl + after
				// But wait, Go's Expand logic handles  groups in repl?
				// regexp.ReplaceAllString does.
				// We need simple first replacement.
				// Actually, VBScript RegExp.Replace uses  syntax too.
				// So we should use re.ExpandString? Or wrap it?
				// Easier approach for Global=False: Use Split/Replace with limit?
				// Go regex doesn't have "ReplaceFirstString".
				// We can do: FindIndex, then if found, replace that range, respecting Expand?
				// Let's assume Global=True behavior for simple ReplaceAll, or just first.
				// For Global=False:
				// "Returns a copy of src, replacing matches of the Regexp with the replacement string repl."
				// We can use Split(str, 2) logic manually?
				// Let's implement simpler logic for now: Global=True -> ReplaceAll. Global=False -> ReplaceFirst.
				// For ReplaceFirst, we can locate index, then apply replacement.
				// But handling  captures in replacement is tricky manually.
				// Let's check if we can use a workaround.
				// Regexp.ReplaceAllStringFunc? No.
				// Let's stick to ReplaceAllString if Global is true.
				// If Global is false, we try to restrict it.
				// Correct way: use FindStringSubmatchIndex, then use re.Expand to generate replacement, then concat.
				if !r.Global {
					// Find first match
					matchIdx := re.FindStringSubmatchIndex(input)
					if matchIdx == nil {
						return input
					}
					// Expand the template
					var result []byte
					result = re.ExpandString(result, repl, input, matchIdx)
					// Concat: prefix + expansion + suffix
					return input[:matchIdx[0]] + string(result) + input[matchIdx[1]:]
				}
				return re.ReplaceAllString(input, repl)
			}
		}
	}
	return nil
}

// --- Match Object ---

type MatchObject struct {
	Value      string
	FirstIndex int
	Length     int
}

func (m *MatchObject) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "value":
		return m.Value
	case "firstindex":
		return m.FirstIndex
	case "length":
		return m.Length
	}
	return nil
}

func (m *MatchObject) SetProperty(name string, value interface{})             {}
func (m *MatchObject) CallMethod(name string, args []interface{}) interface{} { return nil }

// --- Matches Collection ---

type MatchesCollection struct {
	Matches []interface{}
}

func (mc *MatchesCollection) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "count":
		return len(mc.Matches)
	case "item":
		// This is rarely called directly as property, but handled by Engine logic for default property
		// But if Item(i) is called:
		return nil
	}
	return nil
}

func (mc *MatchesCollection) SetProperty(name string, value interface{}) {}
func (mc *MatchesCollection) CallMethod(name string, args []interface{}) interface{} {
	// Item access
	if strings.EqualFold(name, "item") {
		if len(args) > 0 {
			if idx, ok := args[0].(int); ok {
				if idx >= 0 && idx < len(mc.Matches) {
					return mc.Matches[idx]
				}
			}
			// Try converting to int
			idxStr := fmt.Sprintf("%v", args[0])
			var idx int
			if _, err := fmt.Sscanf(idxStr, "%d", &idx); err == nil {
				if idx >= 0 && idx < len(mc.Matches) {
					return mc.Matches[idx]
				}
			}
		}
	}
	return nil
}

func (mc *MatchesCollection) Enumeration() []interface{} {
	return mc.Matches
}
