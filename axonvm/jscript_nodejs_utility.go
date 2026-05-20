/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimaraes - G3pix Ltda
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
	"fmt"
	neturl "net/url"
	"os"
	"path/filepath"
	"runtime"
	"strings"
)

// jsCreatePathObject allocates the Node.js-compatible path module object.
func (vm *VM) jsCreatePathObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 6)
	obj["__js_type"] = NewString("path")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateOSObject allocates the Node.js-compatible os module object.
func (vm *VM) jsCreateOSObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 6)
	obj["__js_type"] = NewString("os")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateQueryStringObject allocates the Node.js-compatible querystring module object.
func (vm *VM) jsCreateQueryStringObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 6)
	obj["__js_type"] = NewString("querystring")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateURLModuleObject allocates the Node.js-compatible url module object.
func (vm *VM) jsCreateURLModuleObject(urlCtor Value, paramsCtor Value) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 8)
	obj["__js_type"] = NewString("url")
	obj["URL"] = urlCtor
	obj["URLSearchParams"] = paramsCtor
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateURLConstructor allocates the global URL constructor function.
func (vm *VM) jsCreateURLConstructor() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 6)
	obj["__js_type"] = NewString("Function")
	obj["__js_ctor"] = NewString("URL")
	obj["name"] = NewString("URL")
	obj["length"] = NewInteger(1)
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)
	return Value{Type: VTJSFunction, Num: objID}
}

// jsCreateURLSearchParamsConstructor allocates the global URLSearchParams constructor function.
func (vm *VM) jsCreateURLSearchParamsConstructor() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 6)
	obj["__js_type"] = NewString("Function")
	obj["__js_ctor"] = NewString("URLSearchParams")
	obj["name"] = NewString("URLSearchParams")
	obj["length"] = NewInteger(1)
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)
	return Value{Type: VTJSFunction, Num: objID}
}

// jsCallPathMethod dispatches path module methods.
func (vm *VM) jsCallPathMethod(methodName string, args []Value) (Value, bool) {
	switch strings.ToLower(methodName) {
	case "join":
		parts := make([]string, 0, len(args))
		for i := 0; i < len(args); i++ {
			parts = append(parts, vm.valueToString(args[i]))
		}
		if len(parts) == 0 {
			return NewString("."), true
		}
		return NewString(filepath.Join(parts...)), true
	case "resolve":
		if len(args) == 0 {
			cwd, err := os.Getwd()
			if err != nil {
				return NewString("."), true
			}
			return NewString(cwd), true
		}
		parts := make([]string, 0, len(args))
		for i := 0; i < len(args); i++ {
			parts = append(parts, vm.valueToString(args[i]))
		}
		resolved, err := filepath.Abs(filepath.Join(parts...))
		if err != nil {
			return NewString(filepath.Join(parts...)), true
		}
		return NewString(resolved), true
	case "basename":
		if len(args) == 0 {
			return NewString(""), true
		}
		base := filepath.Base(vm.valueToString(args[0]))
		if len(args) > 1 {
			ext := vm.valueToString(args[1])
			if ext != "" && strings.HasSuffix(base, ext) {
				base = base[:len(base)-len(ext)]
			}
		}
		return NewString(base), true
	case "normalize":
		if len(args) == 0 {
			return NewString("."), true
		}
		return NewString(filepath.Clean(vm.valueToString(args[0]))), true
	case "extname":
		if len(args) == 0 {
			return NewString(""), true
		}
		return NewString(filepath.Ext(vm.valueToString(args[0]))), true
	}
	return Value{Type: VTJSUndefined}, false
}

// jsCallOSMethod dispatches os module methods.
func (vm *VM) jsCallOSMethod(methodName string, _ []Value) (Value, bool) {
	switch strings.ToLower(methodName) {
	case "arch":
		return NewString(runtime.GOARCH), true
	case "platform":
		return NewString(runtime.GOOS), true
	case "freemem":
		var mem runtime.MemStats
		runtime.ReadMemStats(&mem)
		if mem.Sys > mem.Alloc {
			return NewDouble(float64(mem.Sys - mem.Alloc)), true
		}
		return NewDouble(0), true
	case "cpus":
		cpuCount := runtime.NumCPU()
		if cpuCount < 0 {
			cpuCount = 0
		}
		entries := make([]Value, cpuCount)
		for i := 0; i < cpuCount; i++ {
			timesID := vm.allocJSID()
			timesObj := make(map[string]Value, 8)
			timesObj["__js_type"] = NewString("Object")
			timesObj["user"] = NewDouble(0)
			timesObj["nice"] = NewDouble(0)
			timesObj["sys"] = NewDouble(0)
			timesObj["idle"] = NewDouble(0)
			timesObj["irq"] = NewDouble(0)
			vm.jsObjectItems[timesID] = timesObj
			vm.jsPropertyItems[timesID] = make(map[string]jsPropertyDescriptor, 8)

			cpuID := vm.allocJSID()
			cpuObj := make(map[string]Value, 8)
			cpuObj["__js_type"] = NewString("Object")
			cpuObj["model"] = NewString("Go CPU")
			cpuObj["speed"] = NewInteger(0)
			cpuObj["times"] = Value{Type: VTJSObject, Num: timesID}
			vm.jsObjectItems[cpuID] = cpuObj
			vm.jsPropertyItems[cpuID] = make(map[string]jsPropertyDescriptor, 8)

			entries[i] = Value{Type: VTJSObject, Num: cpuID}
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, entries)), true
	}
	return Value{Type: VTJSUndefined}, false
}

// jsCallQueryStringMethod dispatches querystring module methods.
func (vm *VM) jsCallQueryStringMethod(methodName string, args []Value) (Value, bool) {
	switch strings.ToLower(methodName) {
	case "parse":
		raw := ""
		if len(args) > 0 {
			raw = vm.valueToString(args[0])
		}
		raw = strings.TrimPrefix(raw, "?")
		values, err := neturl.ParseQuery(raw)
		if err != nil {
			values = make(neturl.Values)
		}
		objID := vm.allocJSID()
		obj := make(map[string]Value, len(values)+2)
		obj["__js_type"] = NewString("Object")
		for key, vals := range values {
			if len(vals) > 0 {
				obj[key] = NewString(vals[0])
			} else {
				obj[key] = NewString("")
			}
		}
		vm.jsObjectItems[objID] = obj
		vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, len(obj)+2)
		return Value{Type: VTJSObject, Num: objID}, true
	case "stringify":
		if len(args) == 0 {
			return NewString(""), true
		}
		values := make(neturl.Values)
		source := args[0]
		if source.Type == VTJSObject || source.Type == VTJSFunction {
			keys := vm.jsObjectOwnPropertyNames(source)
			for i := 0; i < len(keys); i++ {
				k := keys[i]
				if strings.HasPrefix(k, "__js_") {
					continue
				}
				v, deferred := vm.jsMemberGet(source, k)
				if deferred {
					continue
				}
				values.Set(k, vm.valueToString(v))
			}
		}
		return NewString(values.Encode()), true
	}
	return Value{Type: VTJSUndefined}, false
}

// jsCallURLModuleMethod dispatches url module methods.
func (vm *VM) jsCallURLModuleMethod(methodName string, args []Value) (Value, bool) {
	switch strings.ToLower(methodName) {
	case "parse":
		return vm.jsConstructURL(args), true
	case "format":
		if len(args) == 0 {
			return NewString(""), true
		}
		if args[0].Type == VTJSObject && vm.jsObjectStringProperty(args[0], "__js_type") == "URL" {
			return NewString(vm.jsObjectStringProperty(args[0], "__js_url_href")), true
		}
		u, err := neturl.Parse(vm.valueToString(args[0]))
		if err != nil {
			return NewString(""), true
		}
		return NewString(u.String()), true
	case "resolve":
		if len(args) == 0 {
			return NewString(""), true
		}
		fromURL, err := neturl.Parse(vm.valueToString(args[0]))
		if err != nil {
			vm.jsThrowTypeError("url.resolve() failed to parse base URL")
			return Value{Type: VTJSUndefined}, true
		}
		toRaw := ""
		if len(args) > 1 {
			toRaw = vm.valueToString(args[1])
		}
		toURL, err := neturl.Parse(toRaw)
		if err != nil {
			vm.jsThrowTypeError("url.resolve() failed to parse target URL")
			return Value{Type: VTJSUndefined}, true
		}
		return NewString(fromURL.ResolveReference(toURL).String()), true
	}
	return Value{Type: VTJSUndefined}, false
}

// jsConstructURL allocates a URL instance using Go's net/url parser.
func (vm *VM) jsConstructURL(args []Value) Value {
	if len(args) == 0 {
		vm.jsThrowTypeError("URL constructor requires at least one argument")
		return Value{Type: VTJSUndefined}
	}

	input := vm.valueToString(args[0])
	var parsed *neturl.URL

	if len(args) > 1 && args[1].Type != VTJSUndefined {
		baseParsed, err := neturl.Parse(vm.valueToString(args[1]))
		if err != nil {
			vm.jsThrowTypeError("Invalid base URL")
			return Value{Type: VTJSUndefined}
		}
		refParsed, err := neturl.Parse(input)
		if err != nil {
			vm.jsThrowTypeError("Invalid URL")
			return Value{Type: VTJSUndefined}
		}
		parsed = baseParsed.ResolveReference(refParsed)
	} else {
		u, err := neturl.Parse(input)
		if err != nil {
			vm.jsThrowTypeError("Invalid URL")
			return Value{Type: VTJSUndefined}
		}
		if u.Scheme == "" && u.Host == "" {
			vm.jsThrowTypeError("Invalid URL")
			return Value{Type: VTJSUndefined}
		}
		parsed = u
	}

	objID := vm.allocJSID()
	obj := make(map[string]Value, 8)
	obj["__js_type"] = NewString("URL")
	obj["__js_url_href"] = NewString(parsed.String())
	if proto := vm.jsGetIntrinsicPrototype("Object"); proto.Type == VTJSObject {
		obj["__js_proto"] = proto
	}
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
	return Value{Type: VTJSObject, Num: objID}
}

// jsConstructURLSearchParams allocates a URLSearchParams instance.
func (vm *VM) jsConstructURLSearchParams(args []Value) Value {
	raw := ""
	if len(args) > 0 {
		source := args[0]
		switch source.Type {
		case VTString:
			raw = source.Str
		case VTJSObject:
			sourceType := vm.jsObjectStringProperty(source, "__js_type")
			if sourceType == "URLSearchParams" {
				raw = vm.jsObjectStringProperty(source, "__js_qs_raw")
			} else {
				values := make(neturl.Values)
				keys := vm.jsObjectOwnPropertyNames(source)
				for i := 0; i < len(keys); i++ {
					k := keys[i]
					if strings.HasPrefix(k, "__js_") {
						continue
					}
					v, deferred := vm.jsMemberGet(source, k)
					if deferred {
						continue
					}
					values.Set(k, vm.valueToString(v))
				}
				raw = values.Encode()
			}
		default:
			raw = vm.valueToString(source)
		}
	}

	raw = strings.TrimPrefix(raw, "?")

	objID := vm.allocJSID()
	obj := make(map[string]Value, 8)
	obj["__js_type"] = NewString("URLSearchParams")
	obj["__js_qs_raw"] = NewString(raw)
	if proto := vm.jsGetIntrinsicPrototype("Object"); proto.Type == VTJSObject {
		obj["__js_proto"] = proto
	}
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
	return Value{Type: VTJSObject, Num: objID}
}

// jsURLFromObject parses and returns the URL stored in a URL instance object.
func (vm *VM) jsURLFromObject(target Value) (*neturl.URL, bool) {
	href := vm.jsObjectStringProperty(target, "__js_url_href")
	if href == "" {
		return nil, false
	}
	parsed, err := neturl.Parse(href)
	if err != nil {
		return nil, false
	}
	return parsed, true
}

// jsSyncURLSearchParamsFromURL keeps URL.searchParams in sync with URL.search.
func (vm *VM) jsSyncURLSearchParamsFromURL(target Value, parsed *neturl.URL) {
	obj := vm.jsObjectItems[target.Num]
	if obj == nil {
		return
	}
	paramsVal, ok := obj["searchParams"]
	if !ok || paramsVal.Type != VTJSObject {
		return
	}
	paramsObj := vm.jsObjectItems[paramsVal.Num]
	if paramsObj == nil {
		return
	}
	paramsObj["__js_qs_raw"] = NewString(parsed.RawQuery)
}

// jsSyncLinkedURLFromSearchParams propagates URLSearchParams mutations to the linked URL object.
func (vm *VM) jsSyncLinkedURLFromSearchParams(target Value, raw string) {
	obj := vm.jsObjectItems[target.Num]
	if obj == nil {
		return
	}
	linked, ok := obj["__js_url_ref"]
	if !ok || linked.Type != VTJSObject {
		return
	}
	linkedObj := vm.jsObjectItems[linked.Num]
	if linkedObj == nil || vm.jsObjectStringProperty(linked, "__js_type") != "URL" {
		return
	}
	u, ok := vm.jsURLFromObject(linked)
	if !ok {
		return
	}
	u.RawQuery = raw
	linkedObj["__js_url_href"] = NewString(u.String())
}

// jsCallURLInstanceMethod dispatches URL instance methods.
func (vm *VM) jsCallURLInstanceMethod(target Value, methodName string, _ []Value) (Value, bool) {
	u, ok := vm.jsURLFromObject(target)
	if !ok {
		return Value{Type: VTJSUndefined}, false
	}
	switch strings.ToLower(methodName) {
	case "tostring", "tojson":
		return NewString(u.String()), true
	}
	return Value{Type: VTJSUndefined}, false
}

// jsParseSearchParamsValues decodes a URLSearchParams raw query into net/url values.
func (vm *VM) jsParseSearchParamsValues(target Value) neturl.Values {
	raw := vm.jsObjectStringProperty(target, "__js_qs_raw")
	values, err := neturl.ParseQuery(raw)
	if err != nil {
		return make(neturl.Values)
	}
	return values
}

// jsStoreSearchParamsValues stores encoded query data and syncs any linked URL.
func (vm *VM) jsStoreSearchParamsValues(target Value, values neturl.Values) {
	encoded := values.Encode()
	obj := vm.jsObjectItems[target.Num]
	if obj != nil {
		obj["__js_qs_raw"] = NewString(encoded)
	}
	vm.jsSyncLinkedURLFromSearchParams(target, encoded)
}

// jsCallURLSearchParamsMethod dispatches URLSearchParams instance methods.
func (vm *VM) jsCallURLSearchParamsMethod(target Value, methodName string, args []Value) (Value, bool) {
	switch strings.ToLower(methodName) {
	case "append":
		if len(args) < 2 {
			vm.jsThrowTypeError("URLSearchParams.append requires name and value")
			return Value{Type: VTJSUndefined}, true
		}
		values := vm.jsParseSearchParamsValues(target)
		values.Add(vm.valueToString(args[0]), vm.valueToString(args[1]))
		vm.jsStoreSearchParamsValues(target, values)
		return Value{Type: VTJSUndefined}, true
	case "set":
		if len(args) < 2 {
			vm.jsThrowTypeError("URLSearchParams.set requires name and value")
			return Value{Type: VTJSUndefined}, true
		}
		values := vm.jsParseSearchParamsValues(target)
		values.Set(vm.valueToString(args[0]), vm.valueToString(args[1]))
		vm.jsStoreSearchParamsValues(target, values)
		return Value{Type: VTJSUndefined}, true
	case "delete":
		if len(args) < 1 {
			vm.jsThrowTypeError("URLSearchParams.delete requires name")
			return Value{Type: VTJSUndefined}, true
		}
		values := vm.jsParseSearchParamsValues(target)
		values.Del(vm.valueToString(args[0]))
		vm.jsStoreSearchParamsValues(target, values)
		return Value{Type: VTJSUndefined}, true
	case "get":
		if len(args) < 1 {
			vm.jsThrowTypeError("URLSearchParams.get requires name")
			return Value{Type: VTJSUndefined}, true
		}
		values := vm.jsParseSearchParamsValues(target)
		key := vm.valueToString(args[0])
		entry := values[key]
		if len(entry) == 0 {
			return Value{Type: VTNull}, true
		}
		return NewString(entry[0]), true
	case "has":
		if len(args) < 1 {
			vm.jsThrowTypeError("URLSearchParams.has requires name")
			return Value{Type: VTJSUndefined}, true
		}
		values := vm.jsParseSearchParamsValues(target)
		_, found := values[vm.valueToString(args[0])]
		return NewBool(found), true
	case "tostring":
		values := vm.jsParseSearchParamsValues(target)
		return NewString(values.Encode()), true
	}
	return Value{Type: VTJSUndefined}, false
}

// jsHandleNodeURLMemberGet handles URL and URLSearchParams property reads.
func (vm *VM) jsHandleNodeURLMemberGet(target Value, member string) (Value, bool) {
	class := vm.jsObjectStringProperty(target, "__js_type")
	if class == "URL" {
		u, ok := vm.jsURLFromObject(target)
		if !ok {
			return Value{Type: VTJSUndefined}, true
		}
		switch strings.ToLower(member) {
		case "href":
			return NewString(u.String()), true
		case "protocol":
			return NewString(u.Scheme + ":"), true
		case "username":
			if u.User == nil {
				return NewString(""), true
			}
			return NewString(u.User.Username()), true
		case "password":
			if u.User == nil {
				return NewString(""), true
			}
			pwd, _ := u.User.Password()
			return NewString(pwd), true
		case "host":
			return NewString(u.Host), true
		case "hostname":
			return NewString(u.Hostname()), true
		case "port":
			return NewString(u.Port()), true
		case "pathname":
			return NewString(u.EscapedPath()), true
		case "search":
			if u.RawQuery == "" {
				return NewString(""), true
			}
			return NewString("?" + u.RawQuery), true
		case "hash":
			if u.Fragment == "" {
				return NewString(""), true
			}
			return NewString("#" + u.Fragment), true
		case "origin":
			if u.Scheme == "" || u.Host == "" {
				return NewString("null"), true
			}
			return NewString(u.Scheme + "://" + u.Host), true
		case "searchparams":
			obj := vm.jsObjectItems[target.Num]
			if obj == nil {
				return Value{Type: VTJSUndefined}, true
			}
			if existing, found := obj["searchParams"]; found && existing.Type == VTJSObject {
				paramsObj := vm.jsObjectItems[existing.Num]
				if paramsObj != nil {
					paramsObj["__js_qs_raw"] = NewString(u.RawQuery)
				}
				return existing, true
			}
			params := vm.jsConstructURLSearchParams([]Value{NewString(u.RawQuery)})
			if params.Type == VTJSObject {
				paramsObj := vm.jsObjectItems[params.Num]
				if paramsObj != nil {
					paramsObj["__js_url_ref"] = target
				}
				obj["searchParams"] = params
			}
			return params, true
		}
	}
	if class == "URLSearchParams" {
		switch strings.ToLower(member) {
		case "size":
			values := vm.jsParseSearchParamsValues(target)
			count := 0
			for _, vals := range values {
				count += len(vals)
			}
			return NewInteger(int64(count)), true
		}
	}
	return Value{Type: VTJSUndefined}, false
}

// jsHandleNodeURLMemberSet handles URL and URLSearchParams property writes.
func (vm *VM) jsHandleNodeURLMemberSet(target Value, member string, val Value) bool {
	class := vm.jsObjectStringProperty(target, "__js_type")
	if class != "URL" {
		return false
	}

	u, ok := vm.jsURLFromObject(target)
	if !ok {
		return true
	}

	switch strings.ToLower(member) {
	case "href":
		parsed, err := neturl.Parse(vm.valueToString(val))
		if err != nil {
			vm.jsThrowTypeError("Invalid URL")
			return true
		}
		u = parsed
	case "protocol":
		raw := vm.valueToString(val)
		raw = strings.TrimSuffix(raw, ":")
		u.Scheme = raw
	case "host":
		u.Host = vm.valueToString(val)
	case "hostname":
		port := u.Port()
		host := vm.valueToString(val)
		if port != "" {
			u.Host = host + ":" + port
		} else {
			u.Host = host
		}
	case "port":
		host := u.Hostname()
		port := vm.valueToString(val)
		if port == "" {
			u.Host = host
		} else {
			u.Host = host + ":" + port
		}
	case "pathname":
		u.Path = vm.valueToString(val)
	case "search":
		raw := vm.valueToString(val)
		raw = strings.TrimPrefix(raw, "?")
		u.RawQuery = raw
	case "hash":
		raw := vm.valueToString(val)
		raw = strings.TrimPrefix(raw, "#")
		u.Fragment = raw
	default:
		return false
	}

	obj := vm.jsObjectItems[target.Num]
	if obj != nil {
		obj["__js_url_href"] = NewString(u.String())
	}
	vm.jsSyncURLSearchParamsFromURL(target, u)
	return true
}

// jsEnsureURLConstructorWithNew validates constructor invocation style for URL classes.
func (vm *VM) jsEnsureURLConstructorWithNew(ctorName string) {
	vm.jsThrowTypeError(fmt.Sprintf("Constructor %s requires 'new'", ctorName))
}
