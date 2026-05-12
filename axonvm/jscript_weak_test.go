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
 */
package axonvm

import (
	"testing"
)

func TestJScriptWeakMapBasics(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var wm = new WeakMap();
		var key = {};
		wm.set(key, "val1");
		Response.Write("has=" + wm.has(key));
		Response.Write(", get=" + wm.get(key));
		wm.delete(key);
		Response.Write(", after_delete=" + wm.has(key));
	</script>`)
	expected := "has=True, get=val1, after_delete=False"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptWeakMapPrimitivesThrow(t *testing.T) {
	_, err := runASPSourceForTestWithErr(t, `<script runat="server" language="JScript">
		var wm = new WeakMap();
		wm.set("string", 1);
	</script>`)
	if err == nil {
		t.Fatal("expected TypeError for primitive key in WeakMap.set, got nil")
	}
}

func TestJScriptWeakMapIsolation(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var wm1 = new WeakMap();
		var wm2 = new WeakMap();
		var key = {};
		wm1.set(key, "wm1");
		wm2.set(key, "wm2");
		Response.Write(wm1.get(key) + "|" + wm2.get(key));
	</script>`)
	if out != "wm1|wm2" {
		t.Errorf("expected 'wm1|wm2', got %q", out)
	}
}

func TestJScriptWeakSetBasics(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var ws = new WeakSet();
		var key = {};
		ws.add(key);
		Response.Write("has=" + ws.has(key));
		ws.delete(key);
		Response.Write(", after_delete=" + ws.has(key));
	</script>`)
	expected := "has=True, after_delete=False"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptWeakMapInitialization(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var k1 = {};
		var k2 = {};
		var wm = new WeakMap([[k1, "v1"], [k2, "v2"]]);
		Response.Write(wm.get(k1) + "|" + wm.get(k2));
	</script>`)
	if out != "v1|v2" {
		t.Errorf("expected 'v1|v2', got %q", out)
	}
}

func TestJScriptWeakCollectionsRequireNew(t *testing.T) {
	_, err := runASPSourceForTestWithErr(t, `<script runat="server" language="JScript">
		WeakMap();
	</script>`)
	if err == nil {
		t.Fatal("expected TypeError when calling WeakMap without new, got nil")
	}

	_, err = runASPSourceForTestWithErr(t, `<script runat="server" language="JScript">
		WeakSet();
	</script>`)
	if err == nil {
		t.Fatal("expected TypeError when calling WeakSet without new, got nil")
	}
}

func TestJScriptWeakMapIterableInitialization(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var k1 = {};
		var k2 = function() {};
		var src = {
			[Symbol.iterator]: function() {
				var index = 0;
				return {
					next: function() {
						index = index + 1;
						if (index === 1) {
							return { value: [k1, "v1"], done: false };
						}
						if (index === 2) {
							return { value: [k2, "v2"], done: false };
						}
						return { value: undefined, done: true };
					}
				};
			}
		};
		var wm = new WeakMap(src);
		Response.Write(wm.get(k1) + "|" + wm.get(k2));
	</script>`)
	if out != "v1|v2" {
		t.Errorf("expected 'v1|v2', got %q", out)
	}
}

func TestJScriptWeakSetIterableInitialization(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var k1 = {};
		var k2 = function() {};
		var src = {
			[Symbol.iterator]: function() {
				var index = 0;
				return {
					next: function() {
						index = index + 1;
						if (index === 1) {
							return { value: k1, done: false };
						}
						if (index === 2) {
							return { value: k2, done: false };
						}
						return { value: undefined, done: true };
					}
				};
			}
		};
		var ws = new WeakSet(src);
		Response.Write(ws.has(k1) + "|" + ws.has(k2));
	</script>`)
	if out != "True|True" {
		t.Errorf("expected 'True|True', got %q", out)
	}
}

func TestJScriptWeakMapDetachedPrototypeMethods(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var wm = new WeakMap();
		var key = {};
		var set = WeakMap.prototype.set;
		var get = WeakMap.prototype.get;
		set.call(wm, key, "ok");
		Response.Write(get.call(wm, key));
	</script>`)
	if out != "ok" {
		t.Errorf("expected 'ok', got %q", out)
	}

	_, err := runASPSourceForTestWithErr(t, `<script runat="server" language="JScript">
		var get = WeakMap.prototype.get;
		get.call({}, {});
	</script>`)
	if err == nil {
		t.Fatal("expected TypeError for WeakMap.prototype.get with incompatible receiver, got nil")
	}
}

func TestJScriptWeakSetDetachedPrototypeMethods(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var ws = new WeakSet();
		var key = {};
		var add = WeakSet.prototype.add;
		var has = WeakSet.prototype.has;
		add.call(ws, key);
		Response.Write(has.call(ws, key));
	</script>`)
	if out != "True" {
		t.Errorf("expected 'True', got %q", out)
	}

	_, err := runASPSourceForTestWithErr(t, `<script runat="server" language="JScript">
		var has = WeakSet.prototype.has;
		has.call({}, {});
	</script>`)
	if err == nil {
		t.Fatal("expected TypeError for WeakSet.prototype.has with incompatible receiver, got nil")
	}
}

func TestJScriptWeakCollectionsObjectToString(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		Response.Write(Object.prototype.toString.call(new WeakMap()));
		Response.Write("|");
		Response.Write(Object.prototype.toString.call(new WeakSet()));
	</script>`)
	if out != "[object WeakMap]|[object WeakSet]" {
		t.Errorf("expected '[object WeakMap]|[object WeakSet]', got %q", out)
	}
}
