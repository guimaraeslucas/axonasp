/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Jeffrey He (@jeffreyheping)
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

// htaAttrRegex matches attribute="value" pairs
var htaAttrRegex = regexp.MustCompile(`(\w+)\s*=\s*"([^"]*)"`)

// ParseHTATag reads a file and extracts <hta:application> attributes.
// Returns nil if no tag is found.
func ParseHTATag(filePath string) *HtaConfig {
	data, err := os.ReadFile(filePath)
	if err != nil {
		return nil
	}

	match := htaTagRegex.FindStringSubmatch(string(data))
	if match == nil || len(match) < 2 {
		return nil
	}

	attrs := match[1]
	cfg := &HtaConfig{}

	pairs := htaAttrRegex.FindAllStringSubmatch(attrs, -1)
	for _, pair := range pairs {
		if len(pair) < 3 {
			continue
		}
		attr := strings.ToLower(pair[1])
		val := pair[2]

		switch attr {
		case "applicationname":
			cfg.ApplicationName = val
		case "border":
			cfg.Border = val
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

// BoolAttr returns true if the attribute is "yes" (case-insensitive).
func (c *HtaConfig) BoolAttr(val string) bool {
	return strings.EqualFold(val, "yes")
}

// StripHTATag removes the <hta:application ... /> tag from HTML content.
func StripHTATag(html string) string {
	return htaTagRegex.ReplaceAllString(html, "")
}

// FindEntryFile searches for an HTA/ASP entry file in appDir.
// Priority: index.hta, default.hta, index.asp, default.asp, index.html, default.html
func FindEntryFile(appDir string) string {
	candidates := []string{
		"index.hta", "default.hta",
		"index.asp", "default.asp",
		"index.html", "default.html",
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
