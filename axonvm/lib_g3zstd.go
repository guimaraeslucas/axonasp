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
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"g3pix.com.br/axonasp/vbscript"
	"github.com/klauspost/compress/zstd"
)

// G3ZSTD exposes klauspost/compress zstd with one ASP-friendly surface.
type G3ZSTD struct {
	vm             *VM
	lastError      string
	defaultLevel   int
	encoder        *zstd.Encoder
	decoder        *zstd.Decoder
	encoderLevelID int
}

// newG3ZSTDObject creates one native G3ZSTD object.
func (vm *VM) newG3ZSTDObject() Value {
	obj := &G3ZSTD{vm: vm, defaultLevel: 3, encoderLevelID: 3}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3zstdItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet routes readable G3ZSTD properties.
func (z *G3ZSTD) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(strings.TrimSpace(propertyName)) {
	case "lasterror":
		return NewString(z.lastError)
	case "level", "compressionlevel":
		return NewInteger(int64(z.defaultLevel))
	}
	return z.DispatchMethod(propertyName, nil)
}

// DispatchMethod routes all G3ZSTD methods.
func (z *G3ZSTD) DispatchMethod(methodName string, args []Value) Value {
	switch strings.ToLower(strings.TrimSpace(methodName)) {
	case "setlevel", "setcompressionlevel":
		if len(args) < 1 {
			return NewBool(false)
		}
		return NewBool(z.setDefaultLevel(int(z.vm.asInt(args[0]))))
	case "compress":
		if len(args) < 1 {
			return NewEmpty()
		}
		level := z.defaultLevel
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
		level := z.defaultLevel
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
		level := z.defaultLevel
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
		if z.encoder != nil {
			z.encoder.Close()
			z.encoder = nil
		}
		if z.decoder != nil {
			z.decoder.Close()
			z.decoder = nil
		}
		z.lastError = ""
		return NewBool(true)
	}
	return NewEmpty()
}

// setDefaultLevel validates and stores one default zstd compression level.
func (z *G3ZSTD) setDefaultLevel(level int) bool {
	if level < -5 || level > 22 {
		z.raiseError("invalid zstd level", fmt.Errorf("level must be between -5 and 22"))
		return false
	}
	z.defaultLevel = level
	z.encoderLevelID = level
	if z.encoder != nil {
		z.encoder.Close()
		z.encoder = nil
	}
	z.lastError = ""
	return true
}

// methodCompress compresses one payload and returns one ASP byte array.
func (z *G3ZSTD) methodCompress(input Value, level int) Value {
	raw, ok := z.inputBytes(input)
	if !ok {
		return NewEmpty()
	}
	enc, ok := z.encoderForLevel(level)
	if !ok {
		return NewEmpty()
	}
	encoded := enc.EncodeAll(raw, nil)
	z.lastError = ""
	return g3zlibBytesToVBArray(encoded)
}

// methodDecompress decompresses one payload and returns one ASP byte array.
func (z *G3ZSTD) methodDecompress(input Value) Value {
	raw, ok := z.inputBytes(input)
	if !ok {
		return NewEmpty()
	}
	dec, ok := z.decoderObject()
	if !ok {
		return NewEmpty()
	}
	decoded, err := dec.DecodeAll(raw, nil)
	if err != nil {
		z.raiseError("zstd decode failed", err)
		return NewEmpty()
	}
	z.lastError = ""
	return g3zlibBytesToVBArray(decoded)
}

// methodDecompressToString decompresses one payload and returns UTF-8 text.
func (z *G3ZSTD) methodDecompressToString(input Value) Value {
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
func (z *G3ZSTD) methodCompressMany(input Value, level int) Value {
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
func (z *G3ZSTD) methodDecompressMany(input Value) Value {
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

// methodCompressFile streams one file through zstd writer to keep memory usage low.
func (z *G3ZSTD) methodCompressFile(sourcePath string, targetPath string, level int) bool {
	source, target, ok := z.resolveFilePair(sourcePath, targetPath)
	if !ok {
		return false
	}

	input, err := os.Open(source)
	if err != nil {
		z.raiseError("zstd compress file open failed", err)
		return false
	}
	defer input.Close()

	if err := os.MkdirAll(filepath.Dir(target), 0755); err != nil {
		z.raiseError("zstd compress mkdir failed", err)
		return false
	}
	output, err := os.Create(target)
	if err != nil {
		z.raiseError("zstd compress file create failed", err)
		return false
	}
	defer output.Close()

	writer, err := zstd.NewWriter(output, zstd.WithEncoderLevel(zstd.EncoderLevelFromZstd(level)))
	if err != nil {
		z.raiseError("zstd compress writer failed", err)
		return false
	}
	if _, err = io.Copy(writer, input); err != nil {
		writer.Close()
		z.raiseError("zstd compress file copy failed", err)
		return false
	}
	if err = writer.Close(); err != nil {
		z.raiseError("zstd compress file close failed", err)
		return false
	}
	z.lastError = ""
	return true
}

// methodDecompressFile streams one zstd file into a plain output file.
func (z *G3ZSTD) methodDecompressFile(sourcePath string, targetPath string) bool {
	source, target, ok := z.resolveFilePair(sourcePath, targetPath)
	if !ok {
		return false
	}

	input, err := os.Open(source)
	if err != nil {
		z.raiseError("zstd decompress file open failed", err)
		return false
	}
	defer input.Close()

	decoder, err := zstd.NewReader(input)
	if err != nil {
		z.raiseError("zstd decompress reader failed", err)
		return false
	}
	defer decoder.Close()

	if err := os.MkdirAll(filepath.Dir(target), 0755); err != nil {
		z.raiseError("zstd decompress mkdir failed", err)
		return false
	}
	output, err := os.Create(target)
	if err != nil {
		z.raiseError("zstd decompress file create failed", err)
		return false
	}
	defer output.Close()

	if _, err = io.Copy(output, decoder); err != nil {
		z.raiseError("zstd decompress file copy failed", err)
		return false
	}
	z.lastError = ""
	return true
}

// encoderForLevel returns one reusable encoder for a specific level.
func (z *G3ZSTD) encoderForLevel(level int) (*zstd.Encoder, bool) {
	if level < -5 || level > 22 {
		z.raiseError("invalid zstd level", fmt.Errorf("level must be between -5 and 22"))
		return nil, false
	}
	if z.encoder != nil && z.encoderLevelID == level {
		return z.encoder, true
	}
	if z.encoder != nil {
		z.encoder.Close()
		z.encoder = nil
	}
	enc, err := zstd.NewWriter(nil, zstd.WithEncoderLevel(zstd.EncoderLevelFromZstd(level)))
	if err != nil {
		z.raiseError("zstd encoder init failed", err)
		return nil, false
	}
	z.encoder = enc
	z.encoderLevelID = level
	return z.encoder, true
}

// decoderObject returns one reusable stateless decoder.
func (z *G3ZSTD) decoderObject() (*zstd.Decoder, bool) {
	if z.decoder != nil {
		return z.decoder, true
	}
	dec, err := zstd.NewReader(nil)
	if err != nil {
		z.raiseError("zstd decoder init failed", err)
		return nil, false
	}
	z.decoder = dec
	return z.decoder, true
}

// inputBytes normalizes one ASP value as bytes using string or VTArray byte semantics.
func (z *G3ZSTD) inputBytes(input Value) ([]byte, bool) {
	if input.Type == VTArray && input.Arr != nil {
		out, ok := g3zlibVBArrayToBytes(z.vm, input)
		if !ok {
			z.raiseError("invalid byte array", fmt.Errorf("array values must be in byte range 0-255"))
			return nil, false
		}
		return out, true
	}
	z.lastError = ""
	return []byte(input.String()), true
}

// resolveFilePair resolves and validates one source-target pair in the FSO sandbox.
func (z *G3ZSTD) resolveFilePair(sourcePath string, targetPath string) (string, string, bool) {
	source, okSource := z.vm.fsoResolvePath(sourcePath)
	target, okTarget := z.vm.fsoResolvePath(targetPath)
	if !okSource || !okTarget {
		z.raiseError("path resolve failed", fmt.Errorf("source or target path is outside the web root sandbox"))
		return "", "", false
	}
	return source, target, true
}

// raiseError stores one error string and raises one VM runtime error.
func (z *G3ZSTD) raiseError(context string, err error) {
	if err == nil {
		return
	}
	message := fmt.Sprintf("%s: %v", context, err)
	z.lastError = message
	z.vm.raise(vbscript.InternalError, message)
}
