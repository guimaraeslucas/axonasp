/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 */
package axonvm

import (
	"testing"
)

func TestJScriptReflectAPI(t *testing.T) {
	tests := []struct {
		code     string
		expected string
	}{
		// 1. Reflect.defineProperty
		{`(function(){
			var obj = {};
			var success = Reflect.defineProperty(obj, "a", {value: 42, writable: true});
			return success + "," + obj.a;
		})()`, "True,42"},

		// 2. Reflect.getOwnPropertyDescriptor
		{`(function(){
			var obj = {a: 1};
			var desc = Reflect.getOwnPropertyDescriptor(obj, "a");
			return desc.value + "," + desc.writable;
		})()`, "1,True"},

		// 3. Reflect.getPrototypeOf
		{`(function(){
			var proto = {p: 1};
			var obj = Object.create(proto);
			return Reflect.getPrototypeOf(obj) === proto;
		})()`, "True"},

		// 4. Reflect.isExtensible
		{`(function(){
			var obj = {};
			var ext1 = Reflect.isExtensible(obj);
			Object.preventExtensions(obj);
			var ext2 = Reflect.isExtensible(obj);
			return ext1 + "," + ext2;
		})()`, "True,False"},

		// 5. Reflect.preventExtensions
		{`(function(){
			var obj = {};
			var success = Reflect.preventExtensions(obj);
			return success + "," + Reflect.isExtensible(obj);
		})()`, "True,False"},

		// 6. Reflect.setPrototypeOf
		{`(function(){
			var obj = {};
			var proto = {a: 1};
			var success = Reflect.setPrototypeOf(obj, proto);
			return success + "," + obj.a;
		})()`, "True,1"},

		// 7. Reflect with Proxy
		{`(function(){
            var target = {a: 1};
            var p = new Proxy(target, {
                defineProperty: function(t, k, d) {
                    if (k === 'b') return false;
                    return Reflect.defineProperty(t, k, d);
                }
            });
            var s1 = Reflect.defineProperty(p, "a", {value: 2});
            var s2 = Reflect.defineProperty(p, "b", {value: 3});
            return s1 + "," + s2 + "," + target.a;
        })()`, "True,False,2"},

		// 8. Object.defineProperty with Proxy
		{`(function(){
            var target = {a: 1};
            var p = new Proxy(target, {
                defineProperty: function(t, k, d) {
                    if (k === 'b') return false;
                    return Reflect.defineProperty(t, k, d);
                }
            });
            Object.defineProperty(p, "a", {value: 2});
            try {
                Object.defineProperty(p, "b", {value: 3});
                return "fail";
            } catch(e) {
                return target.a + "," + e.message;
            }
        })()`, "2,Object.defineProperty failed"},

		// 9. Object.getOwnPropertyDescriptor with Proxy
		{`(function(){
            var target = {a: 1};
            var p = new Proxy(target, {
                getOwnPropertyDescriptor: function(t, k) {
                    return {value: 99, configurable: true};
                }
            });
            var d = Object.getOwnPropertyDescriptor(p, "a");
            return d.value;
        })()`, "99"},

		// 10. Object.getPrototypeOf with Proxy
		{`(function(){
            var target = {};
            var p = new Proxy(target, {
                getPrototypeOf: function(t) {
                    return {mocked: true};
                }
            });
            return Object.getPrototypeOf(p).mocked;
        })()`, "True"},
	}

	for _, tt := range tests {
		out, err := runJScript2(t, jscriptSrc(`Response.Write(`+tt.code+`);`))
		if err != nil {
			t.Errorf("code %q failed: %v", tt.code, err)
			continue
		}
		if out != tt.expected {
			t.Errorf("code %q: expected %q, got %q", tt.code, tt.expected, out)
		}
	}
}
