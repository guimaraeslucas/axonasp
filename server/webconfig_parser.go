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
package server

import (
	"encoding/xml"
	"fmt"
	"net/url"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

// WebConfigError represents a custom error handler configuration
type WebConfigError struct {
	StatusCode   string `xml:"statusCode,attr"`
	Path         string `xml:"path,attr"`
	ResponseMode string `xml:"responseMode,attr"`
}

// WebConfigHTTPErrors contains the httpErrors configuration
type WebConfigHTTPErrors struct {
	Errors []WebConfigError `xml:"error"`
}

// WebConfigSystemWebServer contains the system.webServer configuration
type WebConfigSystemWebServer struct {
	HTTPErrors WebConfigHTTPErrors   `xml:"httpErrors"`
	Rewrite    WebConfigRewrite      `xml:"rewrite"`
	Redirect   WebConfigHTTPRedirect `xml:"httpRedirect"`
}

// WebConfigRewrite contains URL rewrite rules
type WebConfigRewrite struct {
	Rules WebConfigRewriteRules `xml:"rules"`
}

// WebConfigRewriteRules contains multiple rewrite rules
type WebConfigRewriteRules struct {
	Rules []WebConfigRewriteRule `xml:"rule"`
}

// WebConfigRewriteRule represents a single rewrite rule
type WebConfigRewriteRule struct {
	Name           string                     `xml:"name,attr"`
	StopProcessing string                     `xml:"stopProcessing,attr"`
	Match          WebConfigRewriteMatch      `xml:"match"`
	Conditions     WebConfigRewriteConditions `xml:"conditions"`
	Action         WebConfigRewriteAction     `xml:"action"`
}

// WebConfigRewriteConditions contains multiple rewrite conditions
type WebConfigRewriteConditions struct {
	LogicalGrouping string                      `xml:"logicalGrouping,attr"`
	Conditions      []WebConfigRewriteCondition `xml:"add"`
}

// WebConfigRewriteCondition represents a single rewrite condition
type WebConfigRewriteCondition struct {
	Input      string `xml:"input,attr"`
	MatchType  string `xml:"matchType,attr"`
	Pattern    string `xml:"pattern,attr"`
	IgnoreCase string `xml:"ignoreCase,attr"`
	Negate     string `xml:"negate,attr"`
}

// WebConfigRewriteMatch represents the match criteria for a rule
type WebConfigRewriteMatch struct {
	URL        string `xml:"url,attr"`
	IgnoreCase string `xml:"ignoreCase,attr"`
	Negate     string `xml:"negate,attr"`
}

// WebConfigRewriteAction represents the action for a rule
type WebConfigRewriteAction struct {
	Type              string `xml:"type,attr"`
	URL               string `xml:"url,attr"`
	AppendQueryString string `xml:"appendQueryString,attr"`
	RedirectType      string `xml:"redirectType,attr"`
}

// WebConfigHTTPRedirect represents the httpRedirect configuration
type WebConfigHTTPRedirect struct {
	Enabled            string `xml:"enabled,attr"`
	Destination        string `xml:"destination,attr"`
	ExactDestination   string `xml:"exactDestination,attr"`
	ChildOnly          string `xml:"childOnly,attr"`
	HTTPResponseStatus string `xml:"httpResponseStatus,attr"`
}

// WebConfig represents the root web.config structure
type WebConfig struct {
	XMLName         xml.Name                 `xml:"configuration"`
	SystemWebServer WebConfigSystemWebServer `xml:"system.webServer"`
}

// WebConfigParser handles parsing and retrieving web.config settings
type WebConfigParser struct {
	config       *WebConfig
	configPath   string
	rootDir      string
	rewriteRules []RewriteRule
	httpRedirect *HTTPRedirectConfig
}

// RewriteRule is a compiled rewrite rule
type RewriteRule struct {
	Name              string
	StopProcessing    bool
	Pattern           string
	Regex             *regexp.Regexp
	Negate            bool
	ActionType        string
	ActionURL         string
	AppendQueryString bool
	RedirectType      string
	LogicalGrouping   string
	Conditions        []RewriteCondition
}

// RewriteCondition is a compiled rewrite condition
type RewriteCondition struct {
	Input      string
	MatchType  string
	Pattern    string
	Regex      *regexp.Regexp
	IgnoreCase bool
	Negate     bool
}

// RewriteResult contains the result of applying rewrite rules
type RewriteResult struct {
	Applied          bool
	ActionType       string
	Path             string
	RawQuery         string
	RedirectLocation string
	RedirectStatus   int
}

// HTTPRedirectConfig contains parsed httpRedirect settings
type HTTPRedirectConfig struct {
	Enabled          bool
	Destination      string
	ExactDestination bool
	ChildOnly        bool
	StatusCode       int
}

// NewWebConfigParser creates a new web.config parser
func NewWebConfigParser(rootDir string) *WebConfigParser {
	return &WebConfigParser{
		rootDir:    rootDir,
		configPath: filepath.Join(rootDir, "web.config"),
	}
}

// Load parses the web.config file
func (p *WebConfigParser) Load() error {
	// Check if file exists
	if _, err := os.Stat(p.configPath); os.IsNotExist(err) {
		return fmt.Errorf("web.config not found at %s", p.configPath)
	}

	// Read file
	data, err := os.ReadFile(p.configPath)
	if err != nil {
		return fmt.Errorf("failed to read web.config: %w", err)
	}

	// Parse XML
	var config WebConfig
	err = xml.Unmarshal(data, &config)
	if err != nil {
		return fmt.Errorf("failed to parse web.config XML: %w", err)
	}

	p.config = &config
	p.rewriteRules = parseRewriteRules(config.SystemWebServer.Rewrite)
	p.httpRedirect = parseHTTPRedirect(config.SystemWebServer.Redirect)
	return nil
}

// GetErrorHandlerPath returns the path for a specific status code
// Returns empty string if no handler is configured
func (p *WebConfigParser) GetErrorHandlerPath(statusCode int) (string, string) {
	if p.config == nil {
		return "", ""
	}

	statusCodeStr := fmt.Sprintf("%d", statusCode)
	for _, errorConfig := range p.config.SystemWebServer.HTTPErrors.Errors {
		if errorConfig.StatusCode == statusCodeStr {
			// Clean the path - remove leading slash if present
			path := strings.TrimPrefix(errorConfig.Path, "/")
			return path, errorConfig.ResponseMode
		}
	}

	return "", ""
}

// GetFullErrorHandlerPath returns the full file system path for an error handler
func (p *WebConfigParser) GetFullErrorHandlerPath(statusCode int) (string, string) {
	relativePath, responseMode := p.GetErrorHandlerPath(statusCode)
	if relativePath == "" {
		return "", ""
	}

	// Build full path relative to root directory
	fullPath := filepath.Join(p.rootDir, relativePath)
	return fullPath, responseMode
}

// IsLoaded returns true if web.config has been successfully loaded
func (p *WebConfigParser) IsLoaded() bool {
	return p.config != nil
}

// GetRewriteRules returns compiled rewrite rules
func (p *WebConfigParser) GetRewriteRules() []RewriteRule {
	return p.rewriteRules
}

// GetHTTPRedirectConfig returns parsed httpRedirect settings
func (p *WebConfigParser) GetHTTPRedirectConfig() *HTTPRedirectConfig {
	return p.httpRedirect
}

// ApplyRewriteRules applies rewrite rules to the provided path and query string
func (p *WebConfigParser) ApplyRewriteRules(path string, rawQuery string) (RewriteResult, bool) {
	if p.config == nil || len(p.rewriteRules) == 0 {
		return RewriteResult{}, false
	}

	currentPath := strings.TrimPrefix(path, "/")
	currentQuery := rawQuery
	applied := false

	for _, rule := range p.rewriteRules {
		if rule.Regex == nil {
			continue
		}

		matches := rule.Regex.FindStringSubmatch(currentPath)
		matched := matches != nil
		if rule.Negate {
			matched = !matched
		}
		if !matched {
			continue
		}

		// Evaluate conditions
		if !p.evaluateConditions(rule, path, matches) {
			continue
		}

		applied = true
		target := replaceRuleBackRefs(rule.ActionURL, matches)
		targetPath, targetQuery, isAbsolute := splitTargetURL(target)

		switch strings.ToLower(rule.ActionType) {
		case "redirect":
			finalQuery := mergeQueryString(targetQuery, currentQuery, rule.AppendQueryString)
			location := buildLocation(target, targetPath, finalQuery, isAbsolute)
			return RewriteResult{
				Applied:          true,
				ActionType:       "redirect",
				RedirectLocation: location,
				RedirectStatus:   mapRedirectStatus(rule.RedirectType),
			}, true
		case "rewrite":
			if targetPath != "" {
				currentPath = strings.TrimPrefix(targetPath, "/")
			}
			currentQuery = mergeQueryString(targetQuery, currentQuery, rule.AppendQueryString)
		default:
			continue
		}

		if rule.StopProcessing {
			break
		}
	}

	if !applied {
		return RewriteResult{}, false
	}

	return RewriteResult{
		Applied:    true,
		ActionType: "rewrite",
		Path:       "/" + currentPath,
		RawQuery:   currentQuery,
	}, true
}

func (p *WebConfigParser) evaluateConditions(rule RewriteRule, path string, matches []string) bool {
	if len(rule.Conditions) == 0 {
		return true
	}

	logicalGrouping := strings.ToLower(rule.LogicalGrouping)
	if logicalGrouping == "" {
		logicalGrouping = "matchall"
	}

	results := make([]bool, len(rule.Conditions))
	for i, cond := range rule.Conditions {
		input := cond.Input
		// Replace backreferences in input if any
		input = replaceRuleBackRefs(input, matches)

		// Replace common variables
		if input == "{REQUEST_FILENAME}" {
			// Resolve path to physical file
			relPath := strings.TrimPrefix(path, "/")
			input = filepath.Join(p.rootDir, filepath.FromSlash(relPath))
		}

		matched := false
		switch strings.ToLower(cond.MatchType) {
		case "isfile":
			info, err := os.Stat(input)
			matched = err == nil && !info.IsDir()
		case "isdirectory":
			info, err := os.Stat(input)
			matched = err == nil && info.IsDir()
		case "pattern":
			if cond.Regex != nil {
				matched = cond.Regex.MatchString(input)
			}
		}

		if cond.Negate {
			matched = !matched
		}
		results[i] = matched
	}

	if logicalGrouping == "matchany" {
		for _, res := range results {
			if res {
				return true
			}
		}
		return false
	}

	// Default: MatchAll
	for _, res := range results {
		if !res {
			return false
		}
	}
	return true
}

func parseRewriteRules(rewrite WebConfigRewrite) []RewriteRule {
	if len(rewrite.Rules.Rules) == 0 {
		return nil
	}

	rules := make([]RewriteRule, 0, len(rewrite.Rules.Rules))
	for _, rule := range rewrite.Rules.Rules {
		pattern := strings.TrimSpace(rule.Match.URL)
		if pattern == "" {
			continue
		}

		ignoreCase := parseWebConfigBool(rule.Match.IgnoreCase, true)
		re, err := compileRewriteRegex(pattern, ignoreCase)
		if err != nil {
			continue
		}

		actionType := strings.ToLower(strings.TrimSpace(rule.Action.Type))
		if actionType == "" {
			continue
		}

		// Parse conditions
		var conditions []RewriteCondition
		for _, cond := range rule.Conditions.Conditions {
			var condRe *regexp.Regexp
			if strings.ToLower(cond.MatchType) == "pattern" || cond.MatchType == "" {
				condRe, _ = compileRewriteRegex(cond.Pattern, parseWebConfigBool(cond.IgnoreCase, true))
			}

			conditions = append(conditions, RewriteCondition{
				Input:      cond.Input,
				MatchType:  cond.MatchType,
				Pattern:    cond.Pattern,
				Regex:      condRe,
				IgnoreCase: parseWebConfigBool(cond.IgnoreCase, true),
				Negate:     parseWebConfigBool(cond.Negate, false),
			})
		}

		rules = append(rules, RewriteRule{
			Name:              rule.Name,
			StopProcessing:    parseWebConfigBool(rule.StopProcessing, false),
			Pattern:           pattern,
			Regex:             re,
			Negate:            parseWebConfigBool(rule.Match.Negate, false),
			ActionType:        actionType,
			ActionURL:         strings.TrimSpace(rule.Action.URL),
			AppendQueryString: parseWebConfigBool(rule.Action.AppendQueryString, true),
			RedirectType:      strings.TrimSpace(rule.Action.RedirectType),
			LogicalGrouping:   rule.Conditions.LogicalGrouping,
			Conditions:        conditions,
		})
	}

	if len(rules) == 0 {
		return nil
	}

	return rules
}

func parseHTTPRedirect(redirect WebConfigHTTPRedirect) *HTTPRedirectConfig {
	if strings.TrimSpace(redirect.Enabled) == "" && strings.TrimSpace(redirect.Destination) == "" {
		return nil
	}

	config := &HTTPRedirectConfig{
		Enabled:          parseWebConfigBool(redirect.Enabled, false),
		Destination:      strings.TrimSpace(redirect.Destination),
		ExactDestination: parseWebConfigBool(redirect.ExactDestination, false),
		ChildOnly:        parseWebConfigBool(redirect.ChildOnly, false),
		StatusCode:       mapRedirectStatus(redirect.HTTPResponseStatus),
	}

	return config
}

func parseWebConfigBool(value string, defaultValue bool) bool {
	if strings.TrimSpace(value) == "" {
		return defaultValue
	}
	value = strings.ToLower(strings.TrimSpace(value))
	return value == "true" || value == "1" || value == "yes"
}

func compileRewriteRegex(pattern string, ignoreCase bool) (*regexp.Regexp, error) {
	if ignoreCase && !strings.Contains(pattern, "(?i)") {
		pattern = "(?i)" + pattern
	}
	return regexp.Compile(pattern)
}

func replaceRuleBackRefs(input string, matches []string) string {
	if input == "" {
		return input
	}
	backRef := regexp.MustCompile(`(?i)\{r:(\d+)\}`)
	return backRef.ReplaceAllStringFunc(input, func(value string) string {
		parts := backRef.FindStringSubmatch(value)
		if len(parts) != 2 {
			return ""
		}
		index, err := strconv.Atoi(parts[1])
		if err != nil || index < 0 || matches == nil || index >= len(matches) {
			return ""
		}
		return matches[index]
	})
}

func splitTargetURL(target string) (string, string, bool) {
	if target == "" {
		return "", "", false
	}
	if strings.HasPrefix(strings.ToLower(target), "http://") || strings.HasPrefix(strings.ToLower(target), "https://") {
		parsed, err := url.Parse(target)
		if err != nil {
			return target, "", true
		}
		return parsed.Path, parsed.RawQuery, true
	}
	parts := strings.SplitN(target, "?", 2)
	if len(parts) == 2 {
		return parts[0], parts[1], false
	}
	return target, "", false
}

func mergeQueryString(targetQuery string, originalQuery string, appendOriginal bool) string {
	if !appendOriginal || originalQuery == "" {
		return targetQuery
	}
	if targetQuery == "" {
		return originalQuery
	}
	return targetQuery + "&" + originalQuery
}

func buildLocation(originalTarget string, path string, query string, isAbsolute bool) string {
	if isAbsolute {
		parsed, err := url.Parse(originalTarget)
		if err != nil {
			if query == "" {
				return originalTarget
			}
			return originalTarget + "?" + query
		}
		if path != "" {
			parsed.Path = path
		}
		parsed.RawQuery = query
		return parsed.String()
	}
	if query == "" {
		return path
	}
	return path + "?" + query
}

func mapRedirectStatus(redirectType string) int {
	value := strings.ToLower(strings.TrimSpace(redirectType))
	switch value {
	case "permanent":
		return 301
	case "temporary":
		return 307
	case "seeother":
		return 303
	case "found":
		return 302
	default:
		return 302
	}
}
