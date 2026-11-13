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
	"testing"
)

func TestJScriptProxyOperations(t *testing.T) {
	tests := []struct {
		code     string
		expected string
	}{
		// 1. has trap
		{`(function(){
			var p = new Proxy({a: 1}, {
				has: function(target, prop) {
					if (prop === 'b') return true;
					return prop in target;
				}
			});
			var arr = [];
			if (Reflect.has(p, 'a')) arr.push('true'); else arr.push('false');
			if (Reflect.has(p, 'b')) arr.push('true'); else arr.push('false');
			if (Reflect.has(p, 'c')) arr.push('true'); else arr.push('false');
			return arr.join(",");
		})()`, "true,true,false"},

		// 2. deleteProperty trap
		{`(function(){
			return "delete_test_start";
		})()`, "delete_test_start"},

		{`(function(){
			var target = {a: 1, b: 2};
			delete target.a;
			target.b = 3;
			return target.b;
		})()`, "3"},

		{`(function(){
			return Object.keys({a: 1, b: 2}).join(",");
		})()`, "a,b"},

		// 3. ownKeys trap (Object.keys)
		{`(function(){
			var p = new Proxy({a: 1, b: 2}, {
				ownKeys: function(target) {
					return ['a', 'c'];
				}
			});
			return Object.keys(p).join(",");
		})()`, "a"},

		// 4. ownKeys trap (for-in)
		{`(function(){
			try {
				var p = new Proxy({a: 1, b: 2}, {
					ownKeys: function(target) {
						return ['a', 'c'];
					}
				});
				var s = "";
				for (var k in p) s += k + ",";
				return s;
			} catch(e) {
				return "ERROR: " + e.message;
			}
		})()`, "a,"},

		// 5. Proxy.revocable
		{`(function(){
			var r = Proxy.revocable({a: 1}, {});
			var p = r.proxy;
			var val1 = p.a;
			r.revoke();
			try { var val2 = p.a; return "fail"; } catch(e) { return val1 + ",revoked"; }
		})()`, "1,revoked"},
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
