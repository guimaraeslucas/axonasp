/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Jeffrey He (@jeffreyheping), Lucas Guimarães (@guimaraeslucas)
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
package main

import (
	"os"
	"regexp"
	"strings"
)

// HtaConfig holds parsed <hta:application> attributes.
type HtaConfig struct {
	ApplicationName string
	Border          string
	BorderStyle     string
	Caption         string
	Icon            string
	MaximizeButton  string
	MinimizeButton  string
	SingleInstance  string
	WindowState     string
	ShowInTaskbar   string
	Scroll          string
	SysMenu         string
	ContextMenu     string
	InnerBorder     string
	Navigable       string
	ScrollFlat      string
	Selection       string
}

// htaTagRegex matches <hta:application ... /> or <hta:application ...>
var htaTagRegex = regexp.MustCompile(`(?is)<hta:application\s+([^>]*?)/?\s*>`)

// htaAttrRegex matches attribute="value", attribute='value', or attribute=value pairs
var htaAttrRegex = regexp.MustCompile(`([a-zA-Z0-9_-]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))`)

// ParseHTATag reads a file and extracts <hta:application> attributes.
// Returns nil if no tag is found.
func ParseHTATag(filePath string) *HtaConfig {
	data, err := os.ReadFile(filePath)
	if err != nil {
		return nil
	}

	match := htaTagRegex.FindStringSubmatch(string(data))
	if len(match) < 2 {
		return nil
	}

	attrs := match[1]
	cfg := &HtaConfig{}

	pairs := htaAttrRegex.FindAllStringSubmatch(attrs, -1)
	for _, pair := range pairs {
		if len(pair) < 2 {
			continue
		}
		attr := strings.ToLower(pair[1])
		val := ""
		if len(pair) > 2 && pair[2] != "" {
			val = pair[2]
		} else if len(pair) > 3 && pair[3] != "" {
			val = pair[3]
		} else if len(pair) > 4 && pair[4] != "" {
			val = pair[4]
		}

		switch attr {
		case "applicationname":
			cfg.ApplicationName = val
		case "border":
			cfg.Border = val
		case "borderstyle":
			cfg.BorderStyle = val
		case "caption":
			cfg.Caption = val
		case "icon":
			cfg.Icon = val
		case "maximizebutton":
			cfg.MaximizeButton = val
		case "minimizebutton":
			cfg.MinimizeButton = val
		case "singleinstance":
			cfg.SingleInstance = val
		case "windowstate":
			cfg.WindowState = val
		case "showintaskbar":
			cfg.ShowInTaskbar = val
		case "scroll":
			cfg.Scroll = val
		case "sysmenu":
			cfg.SysMenu = val
		case "contextmenu":
			cfg.ContextMenu = val
		case "innerborder":
			cfg.InnerBorder = val
		case "navigable":
			cfg.Navigable = val
		case "scrollflat":
			cfg.ScrollFlat = val
		case "selection":
			cfg.Selection = val
		}
	}

	return cfg
}

// BuildHTAStyleCSS evaluates the parsed HtaConfig and returns the compiled
// <style>...</style> block according to the HTA attribute matrix.
// If no relevant attributes are set to supported values, it returns an empty string.
func (c *HtaConfig) BuildHTAStyleCSS() string {
	if c == nil {
		return ""
	}

	var rules []string

	// Border & BorderStyle evaluation:
	// Priority / mapping per specification matrix:
	// border="dialog" -> border: 3px outset #c0c0c0; box-sizing: border-box;
	// border="thin"   -> border: 1px solid #000; box-sizing: border-box;
	// borderStyle="raised" -> border: 1px outset #c0c0c0; box-sizing: border-box;
	// borderStyle="static" -> border: 1px solid #000; box-sizing: border-box;
	// borderStyle="sunken" -> border: 1px inset #c0c0c0; box-sizing: border-box;
	borderVal := strings.ToLower(strings.TrimSpace(c.Border))
	borderStyleVal := strings.ToLower(strings.TrimSpace(c.BorderStyle))

	var borderCSS string
	switch borderVal {
	case "dialog":
		borderCSS = "border: 3px outset #c0c0c0; box-sizing: border-box;"
	case "thin":
		borderCSS = "border: 1px solid #000; box-sizing: border-box;"
	}

	// If border didn't specify a rule, check borderStyle
	if borderCSS == "" {
		switch borderStyleVal {
		case "raised":
			borderCSS = "border: 1px outset #c0c0c0; box-sizing: border-box;"
		case "static":
			borderCSS = "border: 1px solid #000; box-sizing: border-box;"
		case "sunken":
			borderCSS = "border: 1px inset #c0c0c0; box-sizing: border-box;"
		}
	}

	if borderCSS != "" {
		rules = append(rules, "html { "+borderCSS+"; width: 100%; height: 100%; }")
	}

	// Scroll evaluation:
	// scroll="yes"  -> overflow: scroll !important;
	// scroll="no"   -> overflow: hidden !important;
	// scroll="auto" -> overflow: auto !important;
	scrollVal := strings.ToLower(strings.TrimSpace(c.Scroll))
	switch scrollVal {
	case "yes":
		rules = append(rules, "html { overflow: scroll !important; }")
	case "no":
		rules = append(rules, "html { overflow: hidden !important; }")
	case "auto":
		rules = append(rules, "html { overflow: auto !important; }")
	}

	if len(rules) == 0 {
		return ""
	}

	return "<style>" + strings.Join(rules, " ") + "</style>"
}

// BuildHTAIconHTML evaluates the parsed HtaConfig and returns the <link rel="icon">
// tag if an icon attribute with a valid non-empty string path is present.
func (c *HtaConfig) BuildHTAIconHTML() string {
	if c == nil {
		return ""
	}
	iconPath := strings.TrimSpace(c.Icon)
	if iconPath == "" {
		return ""
	}
	return `<link rel="icon" href="` + iconPath + `">`
}

// BuildHTAHeadInjections combines the icon link and style blocks for injection.
func (c *HtaConfig) BuildHTAHeadInjections() string {
	if c == nil {
		return ""
	}
	var sb strings.Builder
	if iconTag := c.BuildHTAIconHTML(); iconTag != "" {
		sb.WriteString(iconTag)
	}
	if styleTag := c.BuildHTAStyleCSS(); styleTag != "" {
		sb.WriteString(styleTag)
	}
	return sb.String()
}

// BoolAttr returns true if the attribute is "yes" (case-insensitive).
func (c *HtaConfig) BoolAttr(val string) bool {
	return strings.EqualFold(val, "yes")
}

// StripHTATag removes the <hta:application ... /> tag from HTML content.
func StripHTATag(html string) string {
	return htaTagRegex.ReplaceAllString(html, "")
}

// vbScriptBlockRegex matches an entire <script ...>...</script> block where:
//   - The opening tag contains language="vbscript" or type="text/vbscript" (case-insensitive)
//   - The opening tag does NOT contain runat="server" (in any quote form)
//
// In classic HTA, these blocks ran inside the embedded IE VBScript engine.
// AxonHTA serves pages to Chrome, which has no VBScript engine, so these
// blocks must be converted to server-side <% %> delimiters.
var vbScriptBlockRegex = regexp.MustCompile(`(?is)<script(\s[^>]*)?>.*?</script\s*>`)

// isVBScriptClientBlock returns true when the <script> opening-tag attributes
// indicate client-side VBScript (language or type is vbscript and runat is not server).
func isVBScriptClientBlock(attrs string) bool {
	low := strings.ToLower(attrs)
	// Require a VBScript language/type marker.
	isVBS := strings.Contains(low, "language=\"vbscript\"") ||
		strings.Contains(low, "language='vbscript'") ||
		strings.Contains(low, "language=vbscript") ||
		strings.Contains(low, "type=\"text/vbscript\"") ||
		strings.Contains(low, "type='text/vbscript'") ||
		strings.Contains(low, "type=text/vbscript")
	if !isVBS {
		return false
	}
	// Exclude blocks already handled by the lexer as server-side.
	isServer := strings.Contains(low, "runat=\"server\"") ||
		strings.Contains(low, "runat='server'") ||
		strings.Contains(low, "runat=server")
	return !isServer
}

// ConvertVBScriptTagsToASP rewrites client-side VBScript <script> blocks into
// server-side ASP <% %> delimiters so the AxonASP VBScript engine executes them.
//
// Only blocks matching isVBScriptClientBlock are rewritten. All other script
// blocks (JavaScript, VBScript runat=server, etc.) pass through unchanged.
func ConvertVBScriptTagsToASP(html string) string {
	return vbScriptBlockRegex.ReplaceAllStringFunc(html, func(block string) string {
		// Locate the end of the opening tag to split attrs from body.
		gtIdx := strings.IndexByte(block, '>')
		if gtIdx < 0 {
			return block
		}
		openTag := block[:gtIdx+1]

		// Extract just the attribute portion (between <script and >).
		const scriptPrefix = "<script"
		attrStart := len(scriptPrefix)
		if len(openTag) <= attrStart {
			return block
		}
		attrs := openTag[attrStart : len(openTag)-1] // strip leading <script and trailing >

		if !isVBScriptClientBlock(attrs) {
			return block
		}

		// Find the closing tag to extract the script body.
		closingIdx := strings.Index(strings.ToLower(block), "</script")
		if closingIdx < 0 {
			return block
		}
		body := block[gtIdx+1 : closingIdx]
		return "<%" + body + "%>"
	})
}

// FindEntryFile searches for an HTA/ASP entry file in appDir.
// Priority: index.hta, default.hta, index.asp, default.asp, index.html, default.html
func FindEntryFile(appDir string) string {
	candidates := []string{
		"index.hta", "default.hta", "main.hta",
		"index.asp", "default.asp", "main.asp",
		"index.html", "default.html", "main.html",
		"index.htm", "default.htm", "main.htm",
	}
	for _, name := range candidates {
		path := appDir + string(os.PathSeparator) + name
		if _, err := os.Stat(path); err == nil {
			return path
		}
	}
	return ""
}

// MergeWithFlags applies HTA tag values as defaults, overridden by command-line flags.
func (c *HtaConfig) MergeWithFlags(flagTitle string, flagWidth, flagHeight int) (title string, width, height int) {
	// Title: HTA applicationname > "AxonHTA" > flag override
	title = flagTitle
	if c.ApplicationName != "" && flagTitle == "AxonHTA" {
		title = c.ApplicationName
	}

	// Dimensions: only use HTA values if flags were not explicitly set
	// (flag defaults are 1024x768; HTA values would need to be parsed)
	width = flagWidth
	height = flagHeight

	return
}
