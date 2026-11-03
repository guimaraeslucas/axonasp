package vbscript

import (
	"fmt"
	"strings"
)

// ASPRegion describes source structure only. All offsets are UTF-8 byte offsets.
// It deliberately has no execution, HTTP, or include filesystem dependencies.
type ASPRegion struct {
	Kind               string // static, script, expression, directive, include
	Start, End         int
	CodeStart, CodeEnd int
	Language           string
	Path               string
	Virtual            bool
}

// ScanASP reuses the runtime lexer's delimiter and attribute recognition. Unlike
// execution tokens, regions retain untrimmed source positions for static tools.
func ScanASP(text string) ([]ASPRegion, error) {
	l := NewLexer(text)
	var byteOffsets []int
	if !l.asciiOnly {
		byteOffsets = make([]int, 0, l.Length+1)
		for i := range text {
			byteOffsets = append(byteOffsets, i)
		}
		byteOffsets = append(byteOffsets, len(text))
	}
	bytePos := func(i int) int {
		if l.asciiOnly {
			return i
		}
		return byteOffsets[i]
	}
	var regions []ASPRegion
	staticStart := 0
	lang := detectDefaultASPLanguage(text)
	for l.Index < l.Length {
		start := l.Index
		r := ASPRegion{}
		if l.getChar(start) == '<' && l.getChar(start+1) == '%' {
			codeStart := start + 2
			probe := codeStart
			for strings.ContainsRune(" \t\r\n", l.getChar(probe)) && probe < l.Length {
				probe++
			}
			r.Kind = "script"
			if l.getChar(probe) == '@' {
				r.Kind = "directive"
				codeStart = probe + 1
			}
			if l.getChar(probe) == '=' {
				r.Kind = "expression"
				codeStart = probe + 1
			}
			end, blockEnd, ok := l.findASPPercentEndFrom(codeStart)
			if !ok {
				return nil, fmt.Errorf("unterminated ASP block at byte %d", bytePos(start))
			}
			if r.Kind == "directive" {
				if value, ok := extractDirectiveLanguageValue(l.sliceString(codeStart, end)); ok {
					lang = strings.ToLower(strings.TrimSpace(value))
				}
			}
			r.Language = lang
			r.CodeStart, r.CodeEnd = bytePos(codeStart), bytePos(end)
			l.Index = blockEnd
		} else if length, ok, language := l.isScriptServerStart(); ok {
			end, blockEnd, found := l.findScriptEndFrom(start + length)
			if !found {
				return nil, fmt.Errorf("unterminated server script at byte %d", bytePos(start))
			}
			r.Kind, r.Language = "script", language
			r.CodeStart, r.CodeEnd = bytePos(start+length), bytePos(end)
			l.Index = blockEnd
		} else if length, ok, path, virtual := l.sourceIncludeStart(); ok {
			r.Kind, r.Path, r.Virtual = "include", path, virtual
			l.Index += length
		} else {
			l.Index++
			continue
		}
		if staticStart < start {
			regions = append(regions, ASPRegion{Kind: "static", Start: bytePos(staticStart), End: bytePos(start)})
		}
		r.Start, r.End = bytePos(start), bytePos(l.Index)
		regions = append(regions, r)
		staticStart = l.Index
	}
	if staticStart < l.Length {
		regions = append(regions, ASPRegion{Kind: "static", Start: bytePos(staticStart), End: len(text)})
	}
	return regions, nil
}

// ParseASPIncludeComment is the textual preprocessor's include recognition,
// shared by runtime preprocessing and source tooling.
func ParseASPIncludeComment(comment string) (path string, virtual bool, ok bool) {
	upper := strings.ToUpper(comment)
	if !strings.Contains(upper, "#INCLUDE") {
		return "", false, false
	}
	kind := "FILE"
	if strings.Contains(upper, "VIRTUAL") {
		kind = "VIRTUAL"
		virtual = true
	}
	i := strings.Index(upper, kind)
	if i < 0 {
		return "", false, false
	}
	rest := comment[i+len(kind):]
	_, rest, ok = strings.Cut(rest, "=")
	if !ok {
		return "", false, false
	}
	rest = strings.TrimSpace(rest)
	if len(rest) == 0 || (rest[0] != '\'' && rest[0] != '"') {
		return "", false, false
	}
	path, _, ok = strings.Cut(rest[1:], string(rest[0]))
	return path, virtual, ok
}

func (l *Lexer) sourceIncludeStart() (int, bool, string, bool) {
	if l.sliceString(l.Index, l.Index+4) != "<!--" {
		return 0, false, "", false
	}
	for i := l.Index + 4; i < l.Length; i++ {
		if l.sliceString(i, i+3) == "-->" {
			path, virtual, ok := ParseASPIncludeComment(l.sliceString(l.Index, i+3))
			return i + 3 - l.Index, ok, path, virtual
		}
	}
	return 0, false, "", false
}
