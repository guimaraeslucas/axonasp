package axonvm

import (
	"testing"
)

func TestJScriptPhase2ArrayFillAndObjectSymbols(t *testing.T) {
	tests := []struct {
		code     string
		expected string
	}{
		{`[1, 2, 3].fill(4).join(",")`, "4,4,4"},
		{`[1, 2, 3].fill(4, 1).join(",")`, "1,4,4"},
		{`[1, 2, 3].fill(4, 1, 2).join(",")`, "1,4,3"},
		{`[1, 2, 3].fill(4, -2, -1).join(",")`, "1,4,3"},
		{`(function(){ var sym1 = Symbol("a"); var sym2 = Symbol("b"); var obj = { [sym1]: 1, [sym2]: 2, c: 3 }; return Object.getOwnPropertySymbols(obj).length; })()`, "2"},
		{`(function(){ var sym = Symbol("test"); var obj = { [sym]: "val" }; return obj[Object.getOwnPropertySymbols(obj)[0]]; })()`, "val"},
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
