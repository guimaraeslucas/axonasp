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
	"encoding/binary"
	"os"
	"strings"
	"unicode/utf16"
)

const (
	adodbTypeBinary = 1
	adodbTypeText   = 2

	adodbStateClosed = 0
	adodbStateOpen   = 1

	adodbLineLF   = 10
	adodbLineCR   = 13
	adodbLineCRLF = -1
)

// adodbStreamNativeObject stores one ADODB.Stream runtime instance.
type adodbStreamNativeObject struct {
	typ           int
	mode          int
	state         int
	position      int64
	size          int64
	charset       string
	lineSeparator int
	buffer        []byte
	lastFilePath  string
	// bomTextMode is set when WriteText writes a 2-byte BOM + iso-8859-1 content
	// (IIS "unicode" default charset behavior). ReadText must skip the BOM and
	// decode as iso-8859-1 rather than as proper UTF-16.
	bomTextMode bool
}

// newADODBStreamObject allocates one ADODB.Stream native object and returns its VM handle.
func (vm *VM) newADODBStreamObject() Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.adodbStreamItems[objID] = &adodbStreamNativeObject{
		typ:      adodbTypeText,
		mode:     3,
		state:    adodbStateClosed,
		position: 0,
		size:     0,
		// IIS ADODB.Stream defaults to "unicode" (UTF-16 LE) for text-mode streams.
		charset:       "unicode",
		lineSeparator: adodbLineCR,
		buffer:        make([]byte, 0),
		lastFilePath:  "",
	}
	return Value{Type: VTNativeObject, Num: objID}
}

// dispatchADODBStreamMethod routes ADODB.Stream method calls for dynamic native objects.
func (vm *VM) dispatchADODBStreamMethod(objID int64, member string, args []Value) (Value, bool) {
	stream, exists := vm.adodbStreamItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}

	switch {
	case strings.EqualFold(member, "Open"):
		stream.state = adodbStateOpen
		stream.position = 0
		stream.buffer = make([]byte, 0)
		stream.size = 0
		stream.lastFilePath = ""
		stream.bomTextMode = false
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "Close"):
		stream.state = adodbStateClosed
		stream.buffer = nil
		stream.position = 0
		stream.size = 0
		stream.bomTextMode = false
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "Read"):
		return vm.adodbRead(stream, args), true
	case strings.EqualFold(member, "ReadText"):
		return vm.adodbReadText(stream, args), true
	case strings.EqualFold(member, "Write"):
		vm.adodbWrite(stream, args)
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "WriteText"):
		vm.adodbWriteText(stream, args)
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "LoadFromFile"):
		vm.adodbLoadFromFile(stream, args)
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "SaveToFile"):
		vm.adodbSaveToFile(stream, args)
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "CopyTo"):
		vm.adodbCopyTo(stream, args)
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "Flush"):
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "SetEOS"):
		vm.adodbSetEOS(stream)
		return Value{Type: VTEmpty}, true
	case strings.EqualFold(member, "SkipLine"):
		vm.adodbSkipLine(stream)
		return Value{Type: VTEmpty}, true
	default:
		return Value{Type: VTEmpty}, true
	}
}

// dispatchADODBStreamPropertyGet resolves ADODB.Stream property reads.
func (vm *VM) dispatchADODBStreamPropertyGet(objID int64, member string) (Value, bool) {
	stream, exists := vm.adodbStreamItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}

	switch {
	case strings.EqualFold(member, "Type"):
		return NewInteger(int64(stream.typ)), true
	case strings.EqualFold(member, "Mode"):
		return NewInteger(int64(stream.mode)), true
	case strings.EqualFold(member, "State"):
		return NewInteger(int64(stream.state)), true
	case strings.EqualFold(member, "Position"):
		return NewInteger(stream.position), true
	case strings.EqualFold(member, "Size"):
		return NewInteger(stream.size), true
	case strings.EqualFold(member, "Charset"):
		return NewString(stream.charset), true
	case strings.EqualFold(member, "LineSeparator"):
		return NewInteger(int64(stream.lineSeparator)), true
	case strings.EqualFold(member, "EOS"):
		return NewBool(stream.position >= stream.size), true
	case strings.EqualFold(member, "ReadText"):
		return vm.adodbReadText(stream, nil), true
	default:
		return Value{Type: VTEmpty}, true
	}
}

// dispatchADODBStreamPropertySet handles ADODB.Stream writable properties.
func (vm *VM) dispatchADODBStreamPropertySet(objID int64, member string, val Value) bool {
	stream, exists := vm.adodbStreamItems[objID]
	if !exists {
		return false
	}

	switch {
	case strings.EqualFold(member, "Type"):
		stream.typ = vm.asInt(val)
	case strings.EqualFold(member, "Mode"):
		stream.mode = vm.asInt(val)
	case strings.EqualFold(member, "Charset"):
		stream.charset = strings.TrimSpace(val.String())
		if stream.charset == "" {
			stream.charset = "utf-8"
		}
	case strings.EqualFold(member, "LineSeparator"):
		stream.lineSeparator = vm.asInt(val)
	case strings.EqualFold(member, "Position"):
		newPosition := int64(vm.asInt(val))
		if newPosition < 0 {
			newPosition = 0
		}
		if newPosition > stream.size {
			newPosition = stream.size
		}
		stream.position = newPosition
	}

	return true
}

// adodbRead returns binary content from the current stream position.
func (vm *VM) adodbRead(stream *adodbStreamNativeObject, args []Value) Value {
	numBytes := int64(-1)
	if len(args) > 0 {
		numBytes = int64(vm.asInt(args[0]))
	}

	if stream.position >= stream.size || len(stream.buffer) == 0 {
		return NewString("")
	}

	if numBytes < 0 {
		data := stream.buffer[stream.position:]
		stream.position = stream.size
		return NewString(bytesToVBByteString(data))
	}

	if stream.position+numBytes > stream.size {
		numBytes = stream.size - stream.position
	}
	if numBytes <= 0 {
		return NewString("")
	}

	data := stream.buffer[stream.position : stream.position+numBytes]
	stream.position += numBytes
	return NewString(bytesToVBByteString(data))
}

// adodbReadText returns text content decoded with the current charset.
func (vm *VM) adodbReadText(stream *adodbStreamNativeObject, args []Value) Value {
	if stream.state != adodbStateOpen {
		return NewString("")
	}

	numChars := int64(-1)
	if len(args) > 0 {
		numChars = int64(vm.asInt(args[0]))
	}

	remaining := stream.size - stream.position
	if remaining <= 0 {
		return NewString("")
	}

	// For streams written via WriteText with a "unicode" charset (BOM-text mode), the first
	// two bytes of the buffer are the BOM (0xFF 0xFE). Advance past them so they are never
	// included in the decoded output.
	if stream.bomTextMode && stream.position < 2 {
		skip := int64(2) - stream.position
		stream.position += skip
		remaining -= skip
		if remaining <= 0 {
			return NewString("")
		}
	}

	bytesToRead := remaining
	if numChars >= 0 && numChars < bytesToRead {
		bytesToRead = numChars
	}

	data := stream.buffer[stream.position : stream.position+bytesToRead]
	stream.position += bytesToRead
	// bomTextMode streams were written as BOM (0xFF 0xFE) + iso-8859-1 content via WriteText.
	// Decode as iso-8859-1 (BOM already skipped above if needed).
	if stream.bomTextMode {
		return NewString(adodbDecodeText(data, "iso-8859-1"))
	}
	return NewString(adodbDecodeText(data, stream.charset))
}

// adodbWrite appends binary content at the current stream position.
func (vm *VM) adodbWrite(stream *adodbStreamNativeObject, args []Value) {
	if stream.state != adodbStateOpen || len(args) < 1 {
		return
	}

	data := vbByteStringToBytes(args[0].String())
	vm.adodbWriteBytes(stream, data)
}

// adodbWriteText writes text content encoded with the current charset.
func (vm *VM) adodbWriteText(stream *adodbStreamNativeObject, args []Value) {
	if stream.state != adodbStateOpen || len(args) < 1 {
		return
	}

	cs := strings.ToLower(strings.TrimSpace(stream.charset))
	isUnicodeBOM := cs == "unicode" || cs == "utf-16" || cs == "utf-16le"

	// In IIS, ADODB.Stream text mode with its default "unicode" charset prepends a 2-byte
	// UTF-16 LE BOM (0xFF 0xFE) on the very first write to an empty stream, then stores
	// content 1-byte-per-char (iso-8859-1 mapping) rather than as proper UTF-16 pairs.
	// This +2 byte offset makes Stream.Position values align with InStrB byte positions
	// so that Classic ASP code like aspLite can use: Position = InStrB_result + 1 directly.
	if isUnicodeBOM && stream.size == 0 && stream.position == 0 {
		vm.adodbWriteBytes(stream, []byte{0xFF, 0xFE})
		stream.bomTextMode = true
	}

	var data []byte
	if isUnicodeBOM {
		// Content is stored 1-byte-per-char after the BOM (iso-8859-1), not as UTF-16.
		data = adodbEncodeText(args[0].String(), "iso-8859-1")
	} else {
		data = adodbEncodeText(args[0].String(), stream.charset)
	}
	if len(args) > 1 && vm.asInt(args[1]) == 1 {
		switch stream.lineSeparator {
		case adodbLineLF:
			data = append(data, '\n')
		case adodbLineCR:
			data = append(data, '\r')
		default:
			data = append(data, '\r', '\n')
		}
	}

	vm.adodbWriteBytes(stream, data)
}

// adodbWriteBytes writes raw bytes into the stream buffer honoring the current position.
func (vm *VM) adodbWriteBytes(stream *adodbStreamNativeObject, data []byte) {
	if len(data) == 0 {
		return
	}

	if stream.position >= int64(len(stream.buffer)) {
		stream.buffer = append(stream.buffer, data...)
	} else {
		needed := stream.position + int64(len(data))
		if needed > int64(len(stream.buffer)) {
			newBuffer := make([]byte, needed)
			copy(newBuffer, stream.buffer)
			stream.buffer = newBuffer
		}
		copy(stream.buffer[stream.position:], data)
	}

	stream.position += int64(len(data))
	if stream.position > stream.size {
		stream.size = stream.position
	}
}

// adodbLoadFromFile loads one file into the stream buffer inside the server sandbox.
func (vm *VM) adodbLoadFromFile(stream *adodbStreamNativeObject, args []Value) {
	if len(args) < 1 {
		return
	}

	resolvedPath, ok := vm.fsoResolvePath(args[0].String())
	if !ok {
		return
	}

	data, err := os.ReadFile(resolvedPath)
	if err != nil {
		return
	}

	stream.state = adodbStateOpen
	stream.buffer = data
	stream.size = int64(len(data))
	stream.position = 0
	stream.lastFilePath = resolvedPath
}

// adodbSaveToFile writes the current stream buffer to one target file in the sandbox.
func (vm *VM) adodbSaveToFile(stream *adodbStreamNativeObject, args []Value) {
	if stream.state != adodbStateOpen || len(args) < 1 {
		return
	}

	resolvedPath, ok := vm.fsoResolvePath(args[0].String())
	if !ok {
		return
	}

	options := 2
	if len(args) > 1 {
		options = vm.asInt(args[1])
	}

	if options == 1 {
		if _, err := os.Stat(resolvedPath); err == nil {
			return
		}
	}

	_ = os.WriteFile(resolvedPath, stream.buffer, 0644)
}

// adodbCopyTo copies bytes from one stream to another using the current source position.
func (vm *VM) adodbCopyTo(source *adodbStreamNativeObject, args []Value) {
	if len(args) < 1 || args[0].Type != VTNativeObject {
		return
	}

	destination, exists := vm.adodbStreamItems[args[0].Num]
	if !exists {
		return
	}

	numChars := int64(-1)
	if len(args) > 1 {
		numChars = int64(vm.asInt(args[1]))
	}

	if source.position >= source.size || source.position >= int64(len(source.buffer)) {
		return
	}

	if numChars < 0 {
		numChars = source.size - source.position
	}
	if source.position+numChars > source.size {
		numChars = source.size - source.position
	}
	if source.position+numChars > int64(len(source.buffer)) {
		numChars = int64(len(source.buffer)) - source.position
	}
	if numChars <= 0 {
		return
	}

	data := source.buffer[source.position : source.position+numChars]
	vm.adodbWriteBytes(destination, data)
	source.position += numChars
}

// adodbSetEOS truncates the stream at the current position.
func (vm *VM) adodbSetEOS(stream *adodbStreamNativeObject) {
	stream.size = stream.position
	if int64(len(stream.buffer)) > stream.size {
		stream.buffer = stream.buffer[:stream.size]
	}
}

// adodbSkipLine advances the current position to the next line boundary.
func (vm *VM) adodbSkipLine(stream *adodbStreamNativeObject) {
	if stream.state != adodbStateOpen {
		return
	}

	for stream.position < stream.size {
		if stream.buffer[stream.position] == '\n' {
			stream.position++
			return
		}
		if stream.buffer[stream.position] == '\r' {
			stream.position++
			if stream.position < stream.size && stream.buffer[stream.position] == '\n' {
				stream.position++
			}
			return
		}
		stream.position++
	}
}

// adodbDecodeText decodes one byte slice using ADODB.Stream charset semantics.
func adodbDecodeText(data []byte, charset string) string {
	if len(data) == 0 {
		return ""
	}

	if len(data) >= 3 && bytes.Equal(data[:3], []byte{0xEF, 0xBB, 0xBF}) {
		data = data[3:]
	}

	cs := strings.ToLower(strings.TrimSpace(charset))
	switch cs {
	case "iso-8859-1", "windows-1252", "ascii", "us-ascii":
		runes := make([]rune, len(data))
		for i := 0; i < len(data); i++ {
			runes[i] = rune(data[i])
		}
		return string(runes)
	case "unicode", "utf-16", "utf-16le":
		return adodbDecodeUTF16(data, binary.LittleEndian)
	case "utf-16be":
		return adodbDecodeUTF16(data, binary.BigEndian)
	default:
		return string(data)
	}
}

// adodbEncodeText encodes one string using ADODB.Stream charset semantics.
func adodbEncodeText(text string, charset string) []byte {
	cs := strings.ToLower(strings.TrimSpace(charset))
	switch cs {
	case "iso-8859-1", "windows-1252", "ascii", "us-ascii":
		data := make([]byte, 0, len(text))
		for _, r := range text {
			if r <= 0xFF {
				data = append(data, byte(r))
				continue
			}
			data = append(data, '?')
		}
		return data
	case "unicode", "utf-16", "utf-16le":
		runes := []rune(text)
		data := make([]byte, len(runes)*2)
		for i := 0; i < len(runes); i++ {
			binary.LittleEndian.PutUint16(data[i*2:], uint16(runes[i]))
		}
		return data
	case "utf-16be":
		runes := []rune(text)
		data := make([]byte, len(runes)*2)
		for i := 0; i < len(runes); i++ {
			binary.BigEndian.PutUint16(data[i*2:], uint16(runes[i]))
		}
		return data
	default:
		return []byte(text)
	}
}

// adodbDecodeUTF16 decodes one UTF-16 byte sequence using the provided byte order.
func adodbDecodeUTF16(data []byte, order binary.ByteOrder) string {
	if len(data) == 0 {
		return ""
	}
	if len(data)%2 != 0 {
		data = data[:len(data)-1]
	}

	if len(data) >= 2 {
		if data[0] == 0xFF && data[1] == 0xFE {
			data = data[2:]
			order = binary.LittleEndian
		} else if data[0] == 0xFE && data[1] == 0xFF {
			data = data[2:]
			order = binary.BigEndian
		}
	}

	u16 := make([]uint16, len(data)/2)
	for i := range u16 {
		u16[i] = order.Uint16(data[i*2:])
	}
	return string(utf16.Decode(u16))
}
