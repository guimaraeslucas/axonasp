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
package main

import (
	"encoding/xml"
	"fmt"
	"net/http"
	"net/url"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

var webConfigRuleBackRefRE = regexp.MustCompile(`(?i)\{R:(\d+)\}`)

// WebConfigProcessor loads and applies IIS-compatible web.config directives.
type WebConfigProcessor struct {
	rootDir      string
	configPath   string
	rewriteRules []compiledRewriteRule
	httpRedirect *compiledHTTPRedirect
	httpErrors   map[int]WebConfigCustomError
}

// RewriteResult stores the output of a rewrite or redirect decision.
type RewriteResult struct {
	ActionType       string
	Path             string
	RawQuery         string
	RedirectLocation string
	RedirectStatus   int
}

// WebConfigCustomError stores one custom error mapping from web.config.
type WebConfigCustomError struct {
	Path         string
	ResponseMode string
}

type webConfigFile struct {
	XMLName         xml.Name              `xml:"configuration"`
	SystemWebServer webConfigSystemServer `xml:"system.webServer"`
}

type webConfigSystemServer struct {
	HTTPErrors webConfigHTTPErrors `xml:"httpErrors"`
	Rewrite    webConfigRewrite    `xml:"rewrite"`
	Redirect   webConfigHTTPRedir  `xml:"httpRedirect"`
}

type webConfigHTTPErrors struct {
	Errors []webConfigError `xml:"error"`
}

type webConfigError struct {
	StatusCode   string `xml:"statusCode,attr"`
	Path         string `xml:"path,attr"`
	ResponseMode string `xml:"responseMode,attr"`
}

type webConfigRewrite struct {
	Rules webConfigRewriteRules `xml:"rules"`
}

type webConfigRewriteRules struct {
	Rules []webConfigRewriteRule `xml:"rule"`
}

type webConfigRewriteRule struct {
	Name           string                     `xml:"name,attr"`
	StopProcessing string                     `xml:"stopProcessing,attr"`
	Match          webConfigRewriteMatch      `xml:"match"`
	Conditions     webConfigRewriteConditions `xml:"conditions"`
	Action         webConfigRewriteAction     `xml:"action"`
}

type webConfigRewriteMatch struct {
	URL        string `xml:"url,attr"`
	IgnoreCase string `xml:"ignoreCase,attr"`
	Negate     string `xml:"negate,attr"`
}

type webConfigRewriteConditions struct {
	LogicalGrouping string                    `xml:"logicalGrouping,attr"`
	Conditions      []webConfigRewriteCondAdd `xml:"add"`
}

type webConfigRewriteCondAdd struct {
	Input      string `xml:"input,attr"`
	MatchType  string `xml:"matchType,attr"`
	Pattern    string `xml:"pattern,attr"`
	IgnoreCase string `xml:"ignoreCase,attr"`
	Negate     string `xml:"negate,attr"`
}

type webConfigRewriteAction struct {
	Type              string `xml:"type,attr"`
	URL               string `xml:"url,attr"`
	AppendQueryString string `xml:"appendQueryString,attr"`
	RedirectType      string `xml:"redirectType,attr"`
}

type webConfigHTTPRedir struct {
	Enabled            string `xml:"enabled,attr"`
	Destination        string `xml:"destination,attr"`
	ExactDestination   string `xml:"exactDestination,attr"`
	ChildOnly          string `xml:"childOnly,attr"`
	HTTPResponseStatus string `xml:"httpResponseStatus,attr"`
}

type compiledRewriteRule struct {
	StopProcessing    bool
	Regex             *regexp.Regexp
	Negate            bool
	ActionType        string
	ActionURL         string
	AppendQueryString bool
	RedirectType      string
	LogicalGrouping   string
	Conditions        []compiledRewriteCondition
}

type compiledRewriteCondition struct {
	Input      string
	MatchType  string
	Regex      *regexp.Regexp
	Negate     bool
	IgnoreCase bool
}

type compiledHTTPRedirect struct {
	Enabled          bool
	Destination      string
	ExactDestination bool
	ChildOnly        bool
	StatusCode       int
}

// NewWebConfigProcessor parses and compiles web.config from the web root.
func NewWebConfigProcessor(rootDir string) (*WebConfigProcessor, error) {
	configPath := filepath.Join(rootDir, "web.config")
	data, err := os.ReadFile(configPath)
	if err != nil {
		return nil, fmt.Errorf("read web.config: %w", err)
	}

	var parsed webConfigFile
	if err := xml.Unmarshal(data, &parsed); err != nil {
		return nil, fmt.Errorf("parse web.config: %w", err)
	}

	processor := &WebConfigProcessor{
		rootDir:      rootDir,
		configPath:   configPath,
		rewriteRules: compileRewriteRules(parsed.SystemWebServer.Rewrite),
		httpRedirect: compileHTTPRedirect(parsed.SystemWebServer.Redirect),
		httpErrors:   compileCustomErrors(parsed.SystemWebServer.HTTPErrors),
	}

	return processor, nil
}

// Apply evaluates global redirect and rewrite rules against one request path.
func (p *WebConfigProcessor) Apply(requestPath string, rawQuery string) (RewriteResult, bool) {
	if p == nil {
		return RewriteResult{}, false
	}

	if p.httpRedirect != nil && p.httpRedirect.Enabled {
		location, status, ok := p.httpRedirect.apply(requestPath, rawQuery)
		if ok {
			return RewriteResult{ActionType: "redirect", RedirectLocation: location, RedirectStatus: status}, true
		}
	}

	if len(p.rewriteRules) == 0 {
		return RewriteResult{}, false
	}

	currentPath := strings.TrimPrefix(requestPath, "/")
	currentQuery := rawQuery
	appliedRewrite := false

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

		if !p.evaluateConditions(rule, requestPath, matches) {
			continue
		}

		target := replaceRuleBackReferences(rule.ActionURL, matches)
		targetPath, targetQuery, isAbsolute := splitActionTarget(target)
		switch rule.ActionType {
		case "redirect":
			finalQuery := mergeQueryString(targetQuery, currentQuery, rule.AppendQueryString)
			location := buildActionLocation(target, targetPath, finalQuery, isAbsolute)
			return RewriteResult{
				ActionType:       "redirect",
				RedirectLocation: location,
				RedirectStatus:   mapRedirectStatus(rule.RedirectType),
			}, true
		case "rewrite":
			if targetPath != "" {
				currentPath = strings.TrimPrefix(targetPath, "/")
			}
			currentQuery = mergeQueryString(targetQuery, currentQuery, rule.AppendQueryString)
			appliedRewrite = true
		default:
			continue
		}

		if rule.StopProcessing {
			break
		}
	}

	if !appliedRewrite {
		return RewriteResult{}, false
	}

	return RewriteResult{ActionType: "rewrite", Path: "/" + currentPath, RawQuery: currentQuery}, true
}

// GetCustomError resolves one custom httpErrors mapping by status code.
func (p *WebConfigProcessor) GetCustomError(statusCode int) (WebConfigCustomError, bool) {
	if p == nil || len(p.httpErrors) == 0 {
		return WebConfigCustomError{}, false
	}
	entry, ok := p.httpErrors[statusCode]
	return entry, ok
}

func (p *WebConfigProcessor) evaluateConditions(rule compiledRewriteRule, requestPath string, matches []string) bool {
	if len(rule.Conditions) == 0 {
		return true
	}

	logicalGrouping := strings.ToLower(strings.TrimSpace(rule.LogicalGrouping))
	if logicalGrouping == "" {
		logicalGrouping = "matchall"
	}

	results := make([]bool, len(rule.Conditions))
	for i, condition := range rule.Conditions {
		input := replaceRuleBackReferences(condition.Input, matches)
		input = p.resolveConditionInput(input, requestPath)

		matched := false
		switch strings.ToLower(condition.MatchType) {
		case "isfile":
			info, err := os.Stat(input)
			matched = err == nil && !info.IsDir()
		case "isdirectory":
			info, err := os.Stat(input)
			matched = err == nil && info.IsDir()
		case "pattern", "":
			if condition.Regex != nil {
				matched = condition.Regex.MatchString(input)
			}
		}

		if condition.Negate {
			matched = !matched
		}
		results[i] = matched
	}

	if logicalGrouping == "matchany" {
		for _, result := range results {
			if result {
				return true
			}
		}
		return false
	}

	for _, result := range results {
		if !result {
			return false
		}
	}
	return true
}

func (p *WebConfigProcessor) resolveConditionInput(input string, requestPath string) string {
	switch strings.ToUpper(strings.TrimSpace(input)) {
	case "{REQUEST_FILENAME}":
		relPath := strings.TrimPrefix(requestPath, "/")
		return filepath.Join(p.rootDir, filepath.FromSlash(relPath))
	case "{URL}":
		return requestPath
	case "{REQUEST_URI}":
		return requestPath
	default:
		return input
	}
}

func compileCustomErrors(httpErrors webConfigHTTPErrors) map[int]WebConfigCustomError {
	if len(httpErrors.Errors) == 0 {
		return nil
	}
	result := make(map[int]WebConfigCustomError, len(httpErrors.Errors))
	for _, errCfg := range httpErrors.Errors {
		statusCode, err := strconv.Atoi(strings.TrimSpace(errCfg.StatusCode))
		if err != nil || statusCode <= 0 {
			continue
		}
		targetPath := strings.TrimSpace(errCfg.Path)
		if targetPath == "" {
			continue
		}
		result[statusCode] = WebConfigCustomError{Path: targetPath, ResponseMode: strings.TrimSpace(errCfg.ResponseMode)}
	}
	if len(result) == 0 {
		return nil
	}
	return result
}

func compileRewriteRules(rewrite webConfigRewrite) []compiledRewriteRule {
	if len(rewrite.Rules.Rules) == 0 {
		return nil
	}

	rules := make([]compiledRewriteRule, 0, len(rewrite.Rules.Rules))
	for _, candidate := range rewrite.Rules.Rules {
		pattern := strings.TrimSpace(candidate.Match.URL)
		if pattern == "" {
			continue
		}

		re, err := compileRewriteRegex(pattern, parseWebConfigBool(candidate.Match.IgnoreCase, true))
		if err != nil {
			continue
		}

		actionType := strings.ToLower(strings.TrimSpace(candidate.Action.Type))
		if actionType == "" {
			continue
		}

		conditions := make([]compiledRewriteCondition, 0, len(candidate.Conditions.Conditions))
		for _, rawCondition := range candidate.Conditions.Conditions {
			matchType := strings.ToLower(strings.TrimSpace(rawCondition.MatchType))
			if matchType == "" {
				matchType = "pattern"
			}
			ignoreCase := parseWebConfigBool(rawCondition.IgnoreCase, true)
			var condRegex *regexp.Regexp
			if matchType == "pattern" {
				condRegex, _ = compileRewriteRegex(strings.TrimSpace(rawCondition.Pattern), ignoreCase)
			}
			conditions = append(conditions, compiledRewriteCondition{
				Input:      rawCondition.Input,
				MatchType:  matchType,
				Regex:      condRegex,
				IgnoreCase: ignoreCase,
				Negate:     parseWebConfigBool(rawCondition.Negate, false),
			})
		}

		rules = append(rules, compiledRewriteRule{
			StopProcessing:    parseWebConfigBool(candidate.StopProcessing, false),
			Regex:             re,
			Negate:            parseWebConfigBool(candidate.Match.Negate, false),
			ActionType:        actionType,
			ActionURL:         strings.TrimSpace(candidate.Action.URL),
			AppendQueryString: parseWebConfigBool(candidate.Action.AppendQueryString, true),
			RedirectType:      strings.TrimSpace(candidate.Action.RedirectType),
			LogicalGrouping:   strings.TrimSpace(candidate.Conditions.LogicalGrouping),
			Conditions:        conditions,
		})
	}
	if len(rules) == 0 {
		return nil
	}
	return rules
}

func compileHTTPRedirect(redirect webConfigHTTPRedir) *compiledHTTPRedirect {
	if strings.TrimSpace(redirect.Enabled) == "" && strings.TrimSpace(redirect.Destination) == "" {
		return nil
	}

	return &compiledHTTPRedirect{
		Enabled:          parseWebConfigBool(redirect.Enabled, false),
		Destination:      strings.TrimSpace(redirect.Destination),
		ExactDestination: parseWebConfigBool(redirect.ExactDestination, false),
		ChildOnly:        parseWebConfigBool(redirect.ChildOnly, false),
		StatusCode:       mapRedirectStatus(redirect.HTTPResponseStatus),
	}
}

func (r *compiledHTTPRedirect) apply(requestPath string, rawQuery string) (string, int, bool) {
	if r == nil || !r.Enabled {
		return "", 0, false
	}
	if strings.TrimSpace(r.Destination) == "" {
		return "", 0, false
	}
	if r.ChildOnly && (requestPath == "" || requestPath == "/") {
		return "", 0, false
	}

	destination := strings.TrimSpace(r.Destination)
	if !r.ExactDestination {
		destination = appendPathToDestination(destination, requestPath)
	}
	if rawQuery != "" {
		separator := "?"
		if strings.Contains(destination, "?") {
			separator = "&"
		}
		destination += separator + rawQuery
	}

	statusCode := r.StatusCode
	if statusCode == 0 {
		statusCode = http.StatusFound
	}
	return destination, statusCode, true
}

func appendPathToDestination(destination string, requestPath string) string {
	parsed, err := url.Parse(destination)
	if err != nil {
		base := strings.TrimRight(destination, "/")
		suffix := strings.TrimLeft(requestPath, "/")
		if suffix == "" {
			return base + "/"
		}
		if base == "" {
			return "/" + suffix
		}
		return base + "/" + suffix
	}

	basePath := strings.TrimRight(parsed.Path, "/")
	suffixPath := strings.TrimLeft(requestPath, "/")
	if suffixPath == "" {
		if basePath == "" {
			parsed.Path = "/"
		} else {
			parsed.Path = basePath + "/"
		}
		return parsed.String()
	}
	if basePath == "" {
		parsed.Path = "/" + suffixPath
	} else {
		parsed.Path = basePath + "/" + suffixPath
	}
	return parsed.String()
}

func parseWebConfigBool(value string, defaultValue bool) bool {
	if strings.TrimSpace(value) == "" {
		return defaultValue
	}
	value = strings.ToLower(strings.TrimSpace(value))
	return value == "true" || value == "1" || value == "yes"
}

func compileRewriteRegex(pattern string, ignoreCase bool) (*regexp.Regexp, error) {
	if pattern == "" {
		return nil, fmt.Errorf("empty regex pattern")
	}
	if ignoreCase && !strings.Contains(pattern, "(?i)") {
		pattern = "(?i)" + pattern
	}
	return regexp.Compile(pattern)
}

func replaceRuleBackReferences(input string, matches []string) string {
	if input == "" {
		return input
	}
	return webConfigRuleBackRefRE.ReplaceAllStringFunc(input, func(value string) string {
		parts := webConfigRuleBackRefRE.FindStringSubmatch(value)
		if len(parts) != 2 {
			return ""
		}
		index, err := strconv.Atoi(parts[1])
		if err != nil || index < 0 || index >= len(matches) {
			return ""
		}
		return matches[index]
	})
}

func splitActionTarget(target string) (string, string, bool) {
	if target == "" {
		return "", "", false
	}
	lowerTarget := strings.ToLower(target)
	if strings.HasPrefix(lowerTarget, "http://") || strings.HasPrefix(lowerTarget, "https://") {
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

func buildActionLocation(originalTarget string, targetPath string, finalQuery string, isAbsolute bool) string {
	if isAbsolute {
		parsed, err := url.Parse(originalTarget)
		if err != nil {
			if finalQuery == "" {
				return originalTarget
			}
			return originalTarget + "?" + finalQuery
		}
		if targetPath != "" {
			parsed.Path = targetPath
		}
		parsed.RawQuery = finalQuery
		return parsed.String()
	}
	if finalQuery == "" {
		return targetPath
	}
	return targetPath + "?" + finalQuery
}

func mapRedirectStatus(redirectType string) int {
	value := strings.ToLower(strings.TrimSpace(redirectType))
	if parsedInt, err := strconv.Atoi(value); err == nil {
		if parsedInt >= 300 && parsedInt <= 399 {
			return parsedInt
		}
	}
	switch value {
	case "permanent":
		return http.StatusMovedPermanently
	case "temporary":
		return http.StatusTemporaryRedirect
	case "seeother":
		return http.StatusSeeOther
	case "found":
		return http.StatusFound
	case "moved":
		return http.StatusMovedPermanently
	default:
		return http.StatusFound
	}
}

// serveWebConfigCustomError executes web.config httpErrors output mode for one status code.
func serveWebConfigCustomError(w http.ResponseWriter, r *http.Request, statusCode int, customError WebConfigCustomError) bool {
	responseMode := strings.ToLower(strings.TrimSpace(customError.ResponseMode))
	if responseMode == "" {
		responseMode = "executeurl"
	}

	target := strings.TrimSpace(customError.Path)
	if target == "" {
		return false
	}

	switch responseMode {
	case "executeurl":
		resolvedPath, ok := resolveCustomErrorFilePath(target)
		if !ok {
			return false
		}
		info, err := os.Stat(resolvedPath)
		if err != nil || info.IsDir() {
			return false
		}
		if isASPExecutionExtension(resolvedPath) {
			executeASPWithStatus(w, r, resolvedPath, statusCode)
			return true
		}
		serveStaticFileWithMIME(newSingleHeaderResponseWriter(w, statusCode), r, resolvedPath)
		return true
	case "file":
		resolvedPath, ok := resolveCustomErrorFilePath(target)
		if !ok {
			return false
		}
		info, err := os.Stat(resolvedPath)
		if err != nil || info.IsDir() {
			return false
		}
		serveStaticFileWithMIME(newSingleHeaderResponseWriter(w, statusCode), r, resolvedPath)
		return true
	case "redirect":
		http.Redirect(w, r, target, http.StatusFound)
		return true
	default:
		return false
	}
}

func resolveCustomErrorFilePath(target string) (string, bool) {
	if strings.HasPrefix(strings.ToLower(target), "http://") || strings.HasPrefix(strings.ToLower(target), "https://") {
		return "", false
	}

	var candidate string
	if filepath.IsAbs(target) {
		candidate = target
	} else {
		trimmed := strings.TrimPrefix(filepath.ToSlash(target), "/")
		candidate = filepath.Join(RootDir, filepath.FromSlash(trimmed))
	}

	absRoot, err := filepath.Abs(RootDir)
	if err != nil {
		return "", false
	}
	absCandidate, err := filepath.Abs(candidate)
	if err != nil {
		return "", false
	}

	rel, err := filepath.Rel(absRoot, absCandidate)
	if err != nil {
		return "", false
	}
	if rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) {
		return "", false
	}

	cleanPath := path.Clean("/" + filepath.ToSlash(rel))
	if strings.HasPrefix(cleanPath, "/../") || cleanPath == "/.." {
		return "", false
	}

	return absCandidate, true
}
