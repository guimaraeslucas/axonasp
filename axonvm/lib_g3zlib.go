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
	"compress/zlib"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"g3pix.com.br/axonasp/vbscript"
)

// G3ZLIB exposes compress/zlib as one VM native object.
type G3ZLIB struct {
	vm        *VM
	lastError string
}

// newG3ZLIBObject creates one native G3ZLIB object.
func (vm *VM) newG3ZLIBObject() Value {
	obj := &G3ZLIB{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3zlibItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet routes readable G3ZLIB properties.
func (z *G3ZLIB) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(strings.TrimSpace(propertyName)) {
	case "lasterror":
		return NewString(z.lastError)
	}
	return z.DispatchMethod(propertyName, nil)
}

// DispatchMethod routes all G3ZLIB methods.
func (z *G3ZLIB) DispatchMethod(methodName string, args []Value) Value {
	switch strings.ToLower(strings.TrimSpace(methodName)) {
	case "compress":
		if len(args) < 1 {
			return NewEmpty()
		}
		level := zlib.DefaultCompression
		if len(args) >= 2 {
			level = int(z.vm.asInt(args[1]))
		}
		return z.methodCompress(args[0], level)
	case "decompress":
		if len(args) < 1 {
			return NewEmpty()
		}
		return z.methodDecompress(args[0])
	case "decompresstext", "decompressstring":
		if len(args) < 1 {
			return NewString("")
		}
		return z.methodDecompressToString(args[0])
	case "compressmany":
		if len(args) < 1 {
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
		}
		level := zlib.DefaultCompression
		if len(args) >= 2 {
			level = int(z.vm.asInt(args[1]))
		}
		return z.methodCompressMany(args[0], level)
	case "decompressmany":
		if len(args) < 1 {
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
		}
		return z.methodDecompressMany(args[0])
	case "compressfile":
		if len(args) < 2 {
			return NewBool(false)
		}
		level := zlib.DefaultCompression
		if len(args) >= 3 {
			level = int(z.vm.asInt(args[2]))
		}
		return NewBool(z.methodCompressFile(args[0].String(), args[1].String(), level))
	case "decompressfile":
		if len(args) < 2 {
			return NewBool(false)
		}
		return NewBool(z.methodDecompressFile(args[0].String(), args[1].String()))
	case "clear", "initialize", "dispose":
		z.lastError = ""
		return NewBool(true)
	}
	return NewEmpty()
}

// methodCompress compresses one input payload and returns one ASP byte array.
func (z *G3ZLIB) methodCompress(input Value, level int) Value {
	raw, ok := z.inputBytes(input)
	if !ok {
		return NewEmpty()
	}

	var out bytes.Buffer
	writer, err := zlib.NewWriterLevel(&out, level)
	if err != nil {
		z.raiseError("zlib compress writer failed", err)
		return NewEmpty()
	}
	if _, err = writer.Write(raw); err != nil {
		writer.Close()
		z.raiseError("zlib compress write failed", err)
		return NewEmpty()
	}
	if err = writer.Close(); err != nil {
		z.raiseError("zlib compress close failed", err)
		return NewEmpty()
	}
	z.lastError = ""
	return g3zlibBytesToVBArray(out.Bytes())
}

// methodDecompress decompresses one payload and returns one ASP byte array.
func (z *G3ZLIB) methodDecompress(input Value) Value {
	raw, ok := z.inputBytes(input)
	if !ok {
		return NewEmpty()
	}
	reader, err := zlib.NewReader(bytes.NewReader(raw))
	if err != nil {
		z.raiseError("zlib decompress open failed", err)
		return NewEmpty()
	}
	defer reader.Close()

	plain, err := io.ReadAll(reader)
	if err != nil {
		z.raiseError("zlib decompress read failed", err)
		return NewEmpty()
	}
	z.lastError = ""
	return g3zlibBytesToVBArray(plain)
}

// methodDecompressToString decompresses one payload and returns UTF-8 text.
func (z *G3ZLIB) methodDecompressToString(input Value) Value {
	decoded := z.methodDecompress(input)
	if decoded.Type != VTArray || decoded.Arr == nil {
		return NewString("")
	}
	bytesValue, ok := g3zlibVBArrayToBytes(z.vm, decoded)
	if !ok {
		return NewString("")
	}
	return NewString(string(bytesValue))
}

// methodCompressMany compresses each item from one ASP array and returns one ASP array.
func (z *G3ZLIB) methodCompressMany(input Value, level int) Value {
	items := g3zlibNormalizeBatchInput(input)
	output := make([]Value, 0, len(items))
	for i := 0; i < len(items); i++ {
		compressed := z.methodCompress(items[i], level)
		if compressed.Type == VTEmpty {
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
		}
		output = append(output, compressed)
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, output)}
}

// methodDecompressMany decompresses each item from one ASP array and returns one ASP array.
func (z *G3ZLIB) methodDecompressMany(input Value) Value {
	items := g3zlibNormalizeBatchInput(input)
	output := make([]Value, 0, len(items))
	for i := 0; i < len(items); i++ {
		decoded := z.methodDecompress(items[i])
		if decoded.Type == VTEmpty {
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
		}
		output = append(output, decoded)
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, output)}
}

// methodCompressFile streams one file through zlib writer to keep memory usage low.
func (z *G3ZLIB) methodCompressFile(sourcePath string, targetPath string, level int) bool {
	source, target, ok := z.resolveFilePair(sourcePath, targetPath)
	if !ok {
		return false
	}

	input, err := os.Open(source)
	if err != nil {
		z.raiseError("zlib compress file open failed", err)
		return false
	}
	defer input.Close()

	if err := os.MkdirAll(filepath.Dir(target), 0755); err != nil {
		z.raiseError("zlib compress mkdir failed", err)
		return false
	}
	output, err := os.Create(target)
	if err != nil {
		z.raiseError("zlib compress file create failed", err)
		return false
	}
	defer output.Close()

	writer, err := zlib.NewWriterLevel(output, level)
	if err != nil {
		z.raiseError("zlib compress file writer failed", err)
		return false
	}
	if _, err = io.Copy(writer, input); err != nil {
		writer.Close()
		z.raiseError("zlib compress file copy failed", err)
		return false
	}
	if err = writer.Close(); err != nil {
		z.raiseError("zlib compress file close failed", err)
		return false
	}
	z.lastError = ""
	return true
}

// methodDecompressFile streams one zlib file into a plain output file.
func (z *G3ZLIB) methodDecompressFile(sourcePath string, targetPath string) bool {
	source, target, ok := z.resolveFilePair(sourcePath, targetPath)
	if !ok {
		return false
	}

	input, err := os.Open(source)
	if err != nil {
		z.raiseError("zlib decompress file open failed", err)
		return false
	}
	defer input.Close()

	reader, err := zlib.NewReader(input)
	if err != nil {
		z.raiseError("zlib decompress file reader failed", err)
		return false
	}
	defer reader.Close()

	if err := os.MkdirAll(filepath.Dir(target), 0755); err != nil {
		z.raiseError("zlib decompress mkdir failed", err)
		return false
	}
	output, err := os.Create(target)
	if err != nil {
		z.raiseError("zlib decompress file create failed", err)
		return false
	}
	defer output.Close()

	if _, err = io.Copy(output, reader); err != nil {
		z.raiseError("zlib decompress file copy failed", err)
		return false
	}
	z.lastError = ""
	return true
}

// inputBytes normalizes one ASP value as bytes using string or VTArray byte semantics.
func (z *G3ZLIB) inputBytes(input Value) ([]byte, bool) {
	if input.Type == VTArray && input.Arr != nil {
		return g3zlibVBArrayToBytes(z.vm, input)
	}
	z.lastError = ""
	return []byte(input.String()), true
}

// resolveFilePair resolves and validates one source-target pair in the FSO sandbox.
func (z *G3ZLIB) resolveFilePair(sourcePath string, targetPath string) (string, string, bool) {
	source, okSource := z.vm.fsoResolvePath(sourcePath)
	target, okTarget := z.vm.fsoResolvePath(targetPath)
	if !okSource || !okTarget {
		z.raiseError("path resolve failed", errors.New("source or target path is outside the web root sandbox"))
		return "", "", false
	}
	return source, target, true
}

// raiseError stores one error string and raises one VM runtime error.
func (z *G3ZLIB) raiseError(context string, err error) {
	if err == nil {
		return
	}
	message := fmt.Sprintf("%s: %v", context, err)
	z.lastError = message
	z.vm.raise(vbscript.InternalError, message)
}

// g3zlibVBArrayToBytes converts one ASP numeric array into one byte slice.
func g3zlibVBArrayToBytes(vm *VM, input Value) ([]byte, bool) {
	if input.Type != VTArray || input.Arr == nil {
		return nil, false
	}
	values := input.Arr.Values
	out := make([]byte, len(values))
	for i := 0; i < len(values); i++ {
		n := vm.asInt(values[i])
		if n < 0 || n > 255 {
			return nil, false
		}
		out[i] = byte(n)
	}
	return out, true
}

// g3zlibBytesToVBArray converts one byte slice into one ASP numeric array.
func g3zlibBytesToVBArray(data []byte) Value {
	if len(data) == 0 {
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
	}
	items := make([]Value, len(data))
	for i := 0; i < len(data); i++ {
		items[i] = NewInteger(int64(data[i]))
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, items)}
}

// g3zlibNormalizeBatchInput treats one scalar as a 1-item batch and arrays as-is.
func g3zlibNormalizeBatchInput(input Value) []Value {
	if input.Type == VTArray && input.Arr != nil {
		return input.Arr.Values
	}
	return []Value{input}
}
