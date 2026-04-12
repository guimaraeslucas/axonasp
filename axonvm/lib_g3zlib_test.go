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
package axonvm

import (
	"bytes"
	"compress/zlib"
	"io"
	"testing"
)

func TestG3ZLIBCompressDecompressArrayRoundTrip(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = NewMockHost()

	objVal := vm.newG3ZLIBObject()
	obj := vm.g3zlibItems[objVal.Num]
	if obj == nil {
		t.Fatal("expected g3zlib native object")
	}

	input := Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{NewInteger(65), NewInteger(66), NewInteger(67), NewInteger(10)})}
	compressed := obj.DispatchMethod("Compress", []Value{input})
	if compressed.Type != VTArray || compressed.Arr == nil {
		t.Fatalf("expected compressed array, got %#v", compressed)
	}
	decoded := obj.DispatchMethod("Decompress", []Value{compressed})
	if decoded.Type != VTArray || decoded.Arr == nil {
		t.Fatalf("expected decompressed array, got %#v", decoded)
	}
	decodedBytes, ok := g3zlibVBArrayToBytes(vm, decoded)
	if !ok {
		t.Fatal("expected byte array conversion")
	}
	if string(decodedBytes) != "ABC\n" {
		t.Fatalf("unexpected roundtrip payload: %q", string(decodedBytes))
	}
}

func TestG3ZLIBCompressMany(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = NewMockHost()

	objVal := vm.newG3ZLIBObject()
	obj := vm.g3zlibItems[objVal.Num]
	if obj == nil {
		t.Fatal("expected g3zlib native object")
	}

	items := Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{NewString("alpha"), NewString("beta")})}
	res := obj.DispatchMethod("CompressMany", []Value{items})
	if res.Type != VTArray || res.Arr == nil {
		t.Fatalf("expected array result, got %#v", res)
	}
	if len(res.Arr.Values) != 2 {
		t.Fatalf("expected 2 compressed items, got %d", len(res.Arr.Values))
	}
}

func TestG3ZLIBCompatibilityWithStdlib(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = NewMockHost()

	objVal := vm.newG3ZLIBObject()
	obj := vm.g3zlibItems[objVal.Num]
	if obj == nil {
		t.Fatal("expected g3zlib native object")
	}

	compressed := obj.DispatchMethod("Compress", []Value{NewString("hello-zlib")})
	if compressed.Type != VTArray || compressed.Arr == nil {
		t.Fatalf("expected compressed array, got %#v", compressed)
	}
	compressedBytes, ok := g3zlibVBArrayToBytes(vm, compressed)
	if !ok {
		t.Fatal("expected byte conversion")
	}
	reader, err := zlib.NewReader(bytes.NewReader(compressedBytes))
	if err != nil {
		t.Fatalf("stdlib zlib reader failed: %v", err)
	}
	plain, err := io.ReadAll(reader)
	reader.Close()
	if err != nil {
		t.Fatalf("stdlib zlib read failed: %v", err)
	}
	if string(plain) != "hello-zlib" {
		t.Fatalf("unexpected payload: %q", string(plain))
	}
}
