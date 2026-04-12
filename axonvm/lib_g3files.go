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
	"bytes"
	"encoding/binary"
	"os"
	"path/filepath"
	"strings"
	"unicode/utf16"
)

// G3Files implements the G3FILES library as a VM native object.
type G3Files struct {
	vm *VM
}

// newG3FilesObject creates a native G3FILES instance.
func (vm *VM) newG3FilesObject() Value {
	obj := &G3Files{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3filesItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet routes property access to method aliases.
func (g *G3Files) DispatchPropertyGet(propertyName string) Value {
	return g.DispatchMethod(propertyName, nil)
}

// DispatchPropertySet keeps compatibility with property assignment syntax.
func (g *G3Files) DispatchPropertySet(_ string, _ []Value) {
}

// DispatchMethod resolves all G3FILES methods.
func (g *G3Files) DispatchMethod(methodName string, args []Value) Value {
	method := strings.ToLower(strings.TrimSpace(methodName))
	switch method {
	case "exists":
		if len(args) < 1 {
			return NewBool(false)
		}
		path, ok := g.resolvePath(args[0].String())
		if !ok {
			return NewBool(false)
		}
		_, err := os.Stat(path)
		return NewBool(err == nil)
	case "read", "readtext":
		if len(args) < 1 {
			return NewString("")
		}
		path, ok := g.resolvePath(args[0].String())
		if !ok {
			return NewString("")
		}
		data, err := os.ReadFile(path)
		if err != nil {
			return NewString("")
		}
		encoding := ""
		if len(args) >= 2 {
			encoding = args[1].String()
		}
		return NewString(g.decodeWithEncoding(data, encoding))
	case "write", "writetext":
		if len(args) < 2 {
			return NewBool(false)
		}
		path, ok := g.resolvePath(args[0].String())
		if !ok {
			return NewBool(false)
		}
		encoding := "utf-8"
		if len(args) >= 3 {
			encoding = args[2].String()
		}
		lineEnding := ""
		if len(args) >= 4 {
			lineEnding = args[3].String()
		}
		includeBOM := false
		if len(args) >= 5 {
			includeBOM = g.vm.asBool(args[4])
		}
		content := g.normalizeLineEnding(args[1].String(), lineEnding)
		encoded := g.encodeWithEncoding(content, encoding, includeBOM)
		if err := os.MkdirAll(filepath.Dir(path), 0755); err != nil {
			return NewBool(false)
		}
		err := os.WriteFile(path, encoded, 0644)
		return NewBool(err == nil)
	case "append", "appendtext":
		if len(args) < 2 {
			return NewBool(false)
		}
		path, ok := g.resolvePath(args[0].String())
		if !ok {
			return NewBool(false)
		}
		encoding := "utf-8"
		if len(args) >= 3 {
			encoding = args[2].String()
		}
		lineEnding := ""
		if len(args) >= 4 {
			lineEnding = args[3].String()
		}
		content := g.normalizeLineEnding(args[1].String(), lineEnding)
		encoded := g.encodeWithEncoding(content, encoding, false)
		if err := os.MkdirAll(filepath.Dir(path), 0755); err != nil {
			return NewBool(false)
		}
		f, err := os.OpenFile(path, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
		if err != nil {
			return NewBool(false)
		}
		_, writeErr := f.Write(encoded)
		closeErr := f.Close()
		return NewBool(writeErr == nil && closeErr == nil)
	case "delete", "remove":
		if len(args) < 1 {
			return NewBool(false)
		}
		path, ok := g.resolvePath(args[0].String())
		if !ok {
			return NewBool(false)
		}
		err := os.Remove(path)
		return NewBool(err == nil)
	case "copy":
		if len(args) < 2 {
			return NewBool(false)
		}
		sourcePath, sourceOK := g.resolvePath(args[0].String())
		destPath, destOK := g.resolvePath(args[1].String())
		if !sourceOK || !destOK {
			return NewBool(false)
		}
		data, err := os.ReadFile(sourcePath)
		if err != nil {
			return NewBool(false)
		}
		err = os.WriteFile(destPath, data, 0644)
		return NewBool(err == nil)
	case "move", "rename":
		if len(args) < 2 {
			return NewBool(false)
		}
		sourcePath, sourceOK := g.resolvePath(args[0].String())
		destPath, destOK := g.resolvePath(args[1].String())
		if !sourceOK || !destOK {
			return NewBool(false)
		}
		err := os.Rename(sourcePath, destPath)
		return NewBool(err == nil)
	case "size":
		if len(args) < 1 {
			return NewInteger(0)
		}
		path, ok := g.resolvePath(args[0].String())
		if !ok {
			return NewInteger(0)
		}
		info, err := os.Stat(path)
		if err != nil || info.IsDir() {
			return NewInteger(0)
		}
		return NewInteger(info.Size())
	case "mkdir", "makedir":
		if len(args) < 1 {
			return NewBool(false)
		}
		path, ok := g.resolvePath(args[0].String())
		if !ok {
			return NewBool(false)
		}
		err := os.MkdirAll(path, 0755)
		return NewBool(err == nil)
	case "list", "listfiles":
		if len(args) < 1 {
			return NewVBArrayFromStrings(nil)
		}
		path, ok := g.resolvePath(args[0].String())
		if !ok {
			return NewVBArrayFromStrings(nil)
		}
		entries, err := os.ReadDir(path)
		if err != nil {
			return NewVBArrayFromStrings(nil)
		}
		names := make([]string, 0, len(entries))
		for _, entry := range entries {
			if !entry.IsDir() {
				names = append(names, entry.Name())
			}
		}
		return NewVBArrayFromStrings(names)
	case "normalizeeol", "normalizelineendings":
		if len(args) < 2 {
			return NewString("")
		}
		return NewString(g.normalizeLineEnding(args[0].String(), args[1].String()))
	case "converttextencoding":
		if len(args) < 3 {
			return NewString("")
		}
		fromEnc := args[1].String()
		toEnc := args[2].String()
		decoded := g.decodeWithEncoding(g.encodeWithEncoding(args[0].String(), fromEnc, false), fromEnc)
		return NewString(g.decodeWithEncoding(g.encodeWithEncoding(decoded, toEnc, false), toEnc))
	case "convertfileencoding":
		if len(args) < 4 {
			return NewBool(false)
		}
		sourcePath, sourceOK := g.resolvePath(args[0].String())
		destPath, destOK := g.resolvePath(args[1].String())
		if !sourceOK || !destOK {
			return NewBool(false)
		}
		srcEnc := args[2].String()
		dstEnc := args[3].String()
		lineEnding := ""
		if len(args) >= 5 {
			lineEnding = args[4].String()
		}
		includeBOM := false
		if len(args) >= 6 {
			includeBOM = g.vm.asBool(args[5])
		}
		data, err := os.ReadFile(sourcePath)
		if err != nil {
			return NewBool(false)
		}
		decoded := g.decodeWithEncoding(data, srcEnc)
		decoded = g.normalizeLineEnding(decoded, lineEnding)
		encoded := g.encodeWithEncoding(decoded, dstEnc, includeBOM)
		if err := os.MkdirAll(filepath.Dir(destPath), 0755); err != nil {
			return NewBool(false)
		}
		err = os.WriteFile(destPath, encoded, 0644)
		return NewBool(err == nil)
	}
	return NewEmpty()
}

// resolvePath maps one user path into the VM sandbox root.
func (g *G3Files) resolvePath(path string) (string, bool) {
	if g == nil || g.vm == nil {
		return "", false
	}
	return g.vm.fsoResolvePath(path)
}

// normalizeEncoding canonicalizes encoding aliases.
func (g *G3Files) normalizeEncoding(encoding string) string {
	enc := strings.ToLower(strings.TrimSpace(encoding))
	switch enc {
	case "", "utf8", "utf-8", "utf-8-bom":
		return "utf-8"
	case "utf16", "utf-16", "utf-16le", "unicode":
		return "utf-16le"
	case "utf-16be":
		return "utf-16be"
	case "latin1", "iso8859-1", "iso-8859-1", "windows-1252":
		return "iso-8859-1"
	case "ascii", "us-ascii":
		return "ascii"
	default:
		return "utf-8"
	}
}

// decodeWithEncoding decodes bytes to string honoring BOM and explicit encoding.
func (g *G3Files) decodeWithEncoding(data []byte, encoding string) string {
	if len(data) >= 3 && bytes.Equal(data[:3], []byte{0xEF, 0xBB, 0xBF}) {
		return string(data[3:])
	}
	if len(data) >= 2 {
		if data[0] == 0xFF && data[1] == 0xFE {
			return g.decodeUTF16(data[2:], binary.LittleEndian)
		}
		if data[0] == 0xFE && data[1] == 0xFF {
			return g.decodeUTF16(data[2:], binary.BigEndian)
		}
	}

	switch g.normalizeEncoding(encoding) {
	case "utf-16le":
		return g.decodeUTF16(data, binary.LittleEndian)
	case "utf-16be":
		return g.decodeUTF16(data, binary.BigEndian)
	case "ascii":
		out := make([]rune, len(data))
		for i, b := range data {
			if b > 0x7F {
				out[i] = '?'
			} else {
				out[i] = rune(b)
			}
		}
		return string(out)
	case "iso-8859-1":
		out := make([]rune, len(data))
		for i, b := range data {
			out[i] = rune(b)
		}
		return string(out)
	default:
		return string(data)
	}
}

// encodeWithEncoding encodes text to bytes for one target encoding.
func (g *G3Files) encodeWithEncoding(text string, encoding string, includeBOM bool) []byte {
	switch g.normalizeEncoding(encoding) {
	case "utf-16le":
		u16 := utf16.Encode([]rune(text))
		encoded := make([]byte, 0, len(u16)*2+2)
		if includeBOM {
			encoded = append(encoded, 0xFF, 0xFE)
		}
		for _, value := range u16 {
			encoded = append(encoded, byte(value), byte(value>>8))
		}
		return encoded
	case "utf-16be":
		u16 := utf16.Encode([]rune(text))
		encoded := make([]byte, 0, len(u16)*2+2)
		if includeBOM {
			encoded = append(encoded, 0xFE, 0xFF)
		}
		for _, value := range u16 {
			encoded = append(encoded, byte(value>>8), byte(value))
		}
		return encoded
	case "ascii":
		runes := []rune(text)
		encoded := make([]byte, len(runes))
		for i, r := range runes {
			if r > 0x7F {
				encoded[i] = '?'
			} else {
				encoded[i] = byte(r)
			}
		}
		return encoded
	case "iso-8859-1":
		runes := []rune(text)
		encoded := make([]byte, len(runes))
		for i, r := range runes {
			if r > 0xFF {
				encoded[i] = '?'
			} else {
				encoded[i] = byte(r)
			}
		}
		return encoded
	default:
		if includeBOM {
			return append([]byte{0xEF, 0xBB, 0xBF}, []byte(text)...)
		}
		return []byte(text)
	}
}

// decodeUTF16 decodes UTF-16 bytes using one byte order.
func (g *G3Files) decodeUTF16(data []byte, order binary.ByteOrder) string {
	if len(data)%2 != 0 {
		data = data[:len(data)-1]
	}
	if len(data) == 0 {
		return ""
	}
	u16 := make([]uint16, len(data)/2)
	for i := range u16 {
		u16[i] = order.Uint16(data[i*2:])
	}
	return string(utf16.Decode(u16))
}

// normalizeLineEnding rewrites line endings to windows, linux, or mac classic style.
func (g *G3Files) normalizeLineEnding(text string, style string) string {
	normalized := strings.ReplaceAll(text, "\r\n", "\n")
	normalized = strings.ReplaceAll(normalized, "\r", "\n")
	switch strings.ToLower(strings.TrimSpace(style)) {
	case "windows", "crlf", "vbcrlf":
		return strings.ReplaceAll(normalized, "\n", "\r\n")
	case "mac", "cr":
		return strings.ReplaceAll(normalized, "\n", "\r")
	case "linux", "unix", "lf":
		return normalized
	default:
		return normalized
	}
}

// NewVBArrayFromStrings creates a VBArray Value from one string slice.
func NewVBArrayFromStrings(items []string) Value {
	if len(items) == 0 {
		return Value{Type: VTArray, Arr: NewVBArray(0, 0)}
	}
	values := make([]Value, len(items))
	for i := range items {
		values[i] = NewString(items[i])
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}
}

// getFullPath resolves one relative path into an absolute path for diagnostics.
func (g *G3Files) getFullPath(path string) string {
	if resolved, ok := g.resolvePath(path); ok {
		return resolved
	}
	fallback, _ := filepath.Abs(path)
	return fallback
}
