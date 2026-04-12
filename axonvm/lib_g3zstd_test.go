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
	"testing"
)

func TestG3ZSTDRoundTrip(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = NewMockHost()

	objVal := vm.newG3ZSTDObject()
	obj := vm.g3zstdItems[objVal.Num]
	if obj == nil {
		t.Fatal("expected g3zstd native object")
	}

	compressed := obj.DispatchMethod("Compress", []Value{NewString("hello-zstd")})
	if compressed.Type != VTArray || compressed.Arr == nil {
		t.Fatalf("expected compressed array, got %#v", compressed)
	}
	decoded := obj.DispatchMethod("DecompressText", []Value{compressed})
	if decoded.Type != VTString {
		t.Fatalf("expected string from DecompressText, got %#v", decoded)
	}
	if decoded.String() != "hello-zstd" {
		t.Fatalf("unexpected roundtrip payload: %q", decoded.String())
	}
}

func TestG3ZSTDCompressMany(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = NewMockHost()

	objVal := vm.newG3ZSTDObject()
	obj := vm.g3zstdItems[objVal.Num]
	if obj == nil {
		t.Fatal("expected g3zstd native object")
	}

	batch := Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{NewString("a"), NewString("b"), NewString("c")})}
	compressedBatch := obj.DispatchMethod("CompressMany", []Value{batch, NewInteger(6)})
	if compressedBatch.Type != VTArray || compressedBatch.Arr == nil {
		t.Fatalf("expected array result, got %#v", compressedBatch)
	}
	if len(compressedBatch.Arr.Values) != 3 {
		t.Fatalf("expected 3 compressed items, got %d", len(compressedBatch.Arr.Values))
	}

	decodedBatch := obj.DispatchMethod("DecompressMany", []Value{compressedBatch})
	if decodedBatch.Type != VTArray || decodedBatch.Arr == nil {
		t.Fatalf("expected array from DecompressMany, got %#v", decodedBatch)
	}
	if len(decodedBatch.Arr.Values) != 3 {
		t.Fatalf("expected 3 decompressed items, got %d", len(decodedBatch.Arr.Values))
	}
}
