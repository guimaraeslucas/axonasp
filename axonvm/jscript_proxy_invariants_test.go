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

// Phase 5 – Proxy/Reflect Invariant Tests
// Validates that every Proxy trap enforces the ECMAScript "Invariants of the
// Essential Internal Methods" as specified in §10.5.

import (
	"strings"
	"testing"
)

// ---------------------------------------------------------------------------
// Subphase 5.1 – Trap Validation Engine
// ---------------------------------------------------------------------------

func TestProxyInvariant_GetNonConfigurableNonWritable(t *testing.T) {
	// The 'get' trap must return the stored value for a non-configurable,
	// non-writable data property on the target.
	code := `
		var target = {};
		Object.defineProperty(target, 'x', { value: 42, writable: false, configurable: false, enumerable: true });
		var p = new Proxy(target, {
			get: function(t, k) { return 99; }
		});
		try {
			var v = p.x;
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected compile error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_GetNonConfigurableNonWritable_SameValue_OK(t *testing.T) {
	// Returning the exact same value is permitted even for non-configurable,
	// non-writable properties.
	code := `
		var target = {};
		Object.defineProperty(target, 'x', { value: 42, writable: false, configurable: false, enumerable: true });
		var p = new Proxy(target, {
			get: function(t, k) { return 42; }
		});
		Response.Write(p.x);
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if out != "42" {
		t.Errorf("expected '42', got %q", out)
	}
}

func TestProxyInvariant_GetAccessorNoGetter_MustBeUndefined(t *testing.T) {
	// Non-configurable accessor with no getter: trap must return undefined.
	code := `
		var target = {};
		Object.defineProperty(target, 'x', { set: function(v){}, configurable: false, enumerable: true });
		var p = new Proxy(target, {
			get: function(t, k) { return 123; }
		});
		try {
			var v = p.x;
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_HasNonConfigurableCannotBeAbsent(t *testing.T) {
	// 'has' trap cannot return false for a non-configurable own property.
	code := `
		var target = {};
		Object.defineProperty(target, 'x', { value: 1, configurable: false });
		var p = new Proxy(target, {
			has: function(t, k) { return false; }
		});
		try {
			var r = Reflect.has(p, 'x');
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_HasNonExtensibleCannotBeAbsent(t *testing.T) {
	// 'has' trap cannot return false for an own property of a non-extensible target.
	code := `
		var target = { x: 1 };
		Object.preventExtensions(target);
		var p = new Proxy(target, {
			has: function(t, k) { return false; }
		});
		try {
			var r = Reflect.has(p, 'x');
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_SetNonConfigurableNonWritable(t *testing.T) {
	// 'set' trap returning true for a non-configurable, non-writable property
	// with a different value is a violation.
	code := `
		var target = {};
		Object.defineProperty(target, 'x', { value: 42, writable: false, configurable: false });
		var p = new Proxy(target, {
			set: function(t, k, v) { return true; }
		});
		try {
			p.x = 99;
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_SetNonConfigurableNonWritable_SameValue_OK(t *testing.T) {
	// Setting to the identical value on a non-configurable, non-writable property
	// is permitted.
	code := `
		var target = {};
		Object.defineProperty(target, 'x', { value: 42, writable: false, configurable: false });
		var p = new Proxy(target, {
			set: function(t, k, v) { return true; }
		});
		p.x = 42;
		Response.Write("ok");
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if out != "ok" {
		t.Errorf("expected 'ok', got %q", out)
	}
}

func TestProxyInvariant_DeleteNonConfigurable(t *testing.T) {
	// 'deleteProperty' trap cannot return true for a non-configurable property.
	code := `
		var target = {};
		Object.defineProperty(target, 'x', { value: 1, configurable: false });
		var p = new Proxy(target, {
			deleteProperty: function(t, k) { return true; }
		});
		try {
			delete p.x;
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_OwnKeysMissingNonConfigurable(t *testing.T) {
	// 'ownKeys' trap result must include all non-configurable own properties.
	code := `
		var target = {};
		Object.defineProperty(target, 'x', { value: 1, configurable: false });
		var p = new Proxy(target, {
			ownKeys: function(t) { return []; }
		});
		try {
			Object.getOwnPropertyNames(p);
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_OwnKeysNonExtensibleExtraKeys(t *testing.T) {
	// For a non-extensible target, 'ownKeys' must not introduce new keys.
	code := `
		var target = { a: 1 };
		Object.preventExtensions(target);
		var p = new Proxy(target, {
			ownKeys: function(t) { return ['a', 'b']; }
		});
		try {
			Object.getOwnPropertyNames(p);
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_DefinePropertyNonExtensible(t *testing.T) {
	// 'defineProperty' trap cannot return true when the target is non-extensible
	// and the property does not exist on the target.
	// Use Reflect.defineProperty so the invariant error is thrown as a catchable
	// JS TypeError (Object.defineProperty wraps errors in a VBScript envelope).
	code := `
		var target = {};
		Object.preventExtensions(target);
		var p = new Proxy(target, {
			defineProperty: function(t, k, desc) { return true; }
		});
		try {
			Reflect.defineProperty(p, 'x', { value: 1, configurable: true });
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_GetOwnPropertyDescriptorHidesNonConfigurable(t *testing.T) {
	// 'getOwnPropertyDescriptor' cannot return undefined for a non-configurable property.
	code := `
		var target = {};
		Object.defineProperty(target, 'x', { value: 1, configurable: false });
		var p = new Proxy(target, {
			getOwnPropertyDescriptor: function(t, k) { return undefined; }
		});
		try {
			Object.getOwnPropertyDescriptor(p, 'x');
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_GetOwnPropertyDescriptorReportsConfigurableForNonConfigurable(t *testing.T) {
	// Trap cannot report configurable:true for a non-configurable target property.
	code := `
		var target = {};
		Object.defineProperty(target, 'x', { value: 1, configurable: false, writable: false });
		var p = new Proxy(target, {
			getOwnPropertyDescriptor: function(t, k) {
				return { value: 1, configurable: true, writable: false, enumerable: false };
			}
		});
		try {
			Object.getOwnPropertyDescriptor(p, 'x');
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

// ---------------------------------------------------------------------------
// Subphase 5.2 – Prototype & Extensibility Safety
// ---------------------------------------------------------------------------

func TestProxyInvariant_GetPrototypeOfNonExtensible(t *testing.T) {
	// 'getPrototypeOf' must return the actual prototype when target is non-extensible.
	code := `
		var proto = {};
		var target = Object.create(proto);
		Object.preventExtensions(target);
		var differentProto = {};
		var p = new Proxy(target, {
			getPrototypeOf: function(t) { return differentProto; }
		});
		try {
			Object.getPrototypeOf(p);
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_GetPrototypeOfNonExtensible_SameProto_OK(t *testing.T) {
	// Returning the same prototype for a non-extensible target is fine.
	code := `
		var proto = {};
		var target = Object.create(proto);
		Object.preventExtensions(target);
		var p = new Proxy(target, {
			getPrototypeOf: function(t) { return Object.getPrototypeOf(t); }
		});
		var result = Object.getPrototypeOf(p);
		Response.Write(result === proto ? "ok" : "mismatch");
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if out != "ok" {
		t.Errorf("expected 'ok', got %q", out)
	}
}

func TestProxyInvariant_SetPrototypeOfNonExtensible(t *testing.T) {
	// Trap returning true while changing the prototype of a non-extensible target
	// is a violation.
	// Use Reflect.setPrototypeOf so the invariant error surfaces as a catchable
	// JS TypeError (Object.setPrototypeOf wraps errors in a VBScript envelope).
	code := `
		var proto = {};
		var target = Object.create(proto);
		Object.preventExtensions(target);
		var newProto = {};
		var p = new Proxy(target, {
			setPrototypeOf: function(t, p) { return true; }
		});
		try {
			Reflect.setPrototypeOf(p, newProto);
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_SetPrototypeOfNonExtensible_SameProto_OK(t *testing.T) {
	// Trap returning true for the same prototype on a non-extensible target is fine.
	code := `
		var proto = {};
		var target = Object.create(proto);
		Object.preventExtensions(target);
		var p = new Proxy(target, {
			setPrototypeOf: function(t, newP) { return true; }
		});
		Object.setPrototypeOf(p, proto);
		Response.Write("ok");
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if out != "ok" {
		t.Errorf("expected 'ok', got %q", out)
	}
}

func TestProxyInvariant_PreventExtensionsTrueButTargetExtensible(t *testing.T) {
	// 'preventExtensions' trap returning true while the target is still extensible
	// is a spec violation.
	code := `
		var target = {};
		var p = new Proxy(target, {
			preventExtensions: function(t) {
				// Intentionally do NOT call Object.preventExtensions(t)
				return true;
			}
		});
		try {
			Object.preventExtensions(p);
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(out, "invariant") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}

func TestProxyInvariant_PreventExtensions_OK(t *testing.T) {
	// A well-behaved 'preventExtensions' trap must call preventExtensions on the
	// target before returning true.
	code := `
		var target = {};
		var p = new Proxy(target, {
			preventExtensions: function(t) {
				Object.preventExtensions(t);
				return true;
			}
		});
		Object.preventExtensions(p);
		Response.Write(Object.isExtensible(target) ? "extensible" : "sealed");
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if out != "sealed" {
		t.Errorf("expected 'sealed', got %q", out)
	}
}

func TestProxyInvariant_IsExtensibleMustMatchTarget(t *testing.T) {
	// Already enforced before Phase 5, but included here as a regression guard.
	code := `
		var target = {};
		Object.preventExtensions(target);
		var p = new Proxy(target, {
			isExtensible: function(t) { return true; }
		});
		try {
			Object.isExtensible(p);
			Response.Write("no error");
		} catch(e) {
			Response.Write(e.message);
		}
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if strings.Contains(out, "no error") {
		t.Errorf("expected invariant violation error, got: %q", out)
	}
}
