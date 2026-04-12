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
	"bufio"
	"encoding/gob"
	"encoding/xml"
	"fmt"
	"math/rand"
	"os"
	"os/user"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"
	"sync"
	"sync/atomic"
	"time"

	"g3pix.com.br/axonasp/axonconfig"
	"g3pix.com.br/axonasp/vbscript"
)

// --- MSWC.AdRotator ---

type G3AdRotator struct {
	vm          *VM
	Border      int
	Clickable   bool
	TargetFrame string
}

func (vm *VM) newG3AdRotatorObject() Value {
	obj := &G3AdRotator{
		vm:        vm,
		Border:    -1, // -1 means use file default
		Clickable: true,
	}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.mswcAdRotatorItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

func (lib *G3AdRotator) DispatchPropertyGet(name string) Value {
	switch strings.ToLower(name) {
	case "border":
		return NewInteger(int64(lib.Border))
	case "clickable":
		return NewBool(lib.Clickable)
	case "targetframe":
		return NewString(lib.TargetFrame)
	}
	return lib.DispatchMethod(name, nil)
}

func (lib *G3AdRotator) DispatchPropertySet(name string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	val := args[0]

	switch strings.ToLower(name) {
	case "border":
		lib.Border = int(lib.vm.asInt(val))
		return true
	case "clickable":
		lib.Clickable = (val.Type == VTBool && val.Num != 0)
		return true
	case "targetframe":
		lib.TargetFrame = val.String()
		return true
	}
	return false
}

func (lib *G3AdRotator) DispatchMethod(name string, args []Value) Value {
	switch strings.ToLower(name) {
	case "getadvertisement":
		if len(args) < 1 {
			return NewString("")
		}
		path := args[0].String()
		return NewString(lib.GetAdvertisement(path))
	}
	return NewEmpty()
}

func (lib *G3AdRotator) GetAdvertisement(scheduleFile string) string {
	if lib.vm.host == nil || lib.vm.host.Server() == nil {
		return ""
	}
	absPath := lib.vm.host.Server().MapPath(scheduleFile)
	file, err := os.Open(absPath)
	if err != nil {
		return fmt.Sprintf("<!-- AdRotator Error: %v -->", err)
	}
	defer file.Close()

	scanner := bufio.NewScanner(file)
	redirect := ""
	width := ""
	height := ""
	border := "0"
	if lib.Border != -1 {
		border = strconv.Itoa(lib.Border)
	}

	for scanner.Scan() {
		line := strings.TrimSpace(scanner.Text())
		if line == "*" {
			break
		}
		parts := strings.SplitN(line, " ", 2)
		if len(parts) == 2 {
			key := strings.ToUpper(parts[0])
			val := strings.TrimSpace(parts[1])
			switch key {
			case "REDIRECT":
				redirect = val
			case "WIDTH":
				width = val
			case "HEIGHT":
				height = val
			case "BORDER":
				if lib.Border == -1 {
					border = val
				}
			}
		}
	}

	type AdEntry struct {
		AdURL       string
		HomeURL     string
		AltText     string
		Impressions int
	}

	var ads []AdEntry
	totalImpressions := 0

	for scanner.Scan() {
		adURL := scanner.Text()
		if !scanner.Scan() {
			break
		}
		homeURL := scanner.Text()
		if !scanner.Scan() {
			break
		}
		altText := scanner.Text()
		if !scanner.Scan() {
			break
		}
		impressionsStr := scanner.Text()
		impressions, _ := strconv.Atoi(strings.TrimSpace(impressionsStr))

		ads = append(ads, AdEntry{
			AdURL:       strings.TrimSpace(adURL),
			HomeURL:     strings.TrimSpace(homeURL),
			AltText:     strings.TrimSpace(altText),
			Impressions: impressions,
		})
		totalImpressions += impressions
	}

	if len(ads) == 0 {
		return ""
	}

	r := rand.New(rand.NewSource(time.Now().UnixNano()))
	target := r.Intn(totalImpressions)
	current := 0
	var selected AdEntry
	for _, ad := range ads {
		current += ad.Impressions
		if target < current {
			selected = ad
			break
		}
	}

	finalURL := selected.HomeURL
	if redirect != "" && selected.HomeURL != "-" && selected.HomeURL != "" {
		sep := "?"
		if strings.Contains(redirect, "?") {
			sep = "&"
		}
		finalURL = fmt.Sprintf("%s%surl=%s&image=%s", redirect, sep, selected.HomeURL, selected.AdURL)
	}

	html := ""
	if lib.Clickable && selected.HomeURL != "-" && selected.HomeURL != "" {
		targetAttr := ""
		if lib.TargetFrame != "" {
			targetAttr = fmt.Sprintf(" target=\"%s\"", lib.TargetFrame)
		}
		html += fmt.Sprintf("<a href=\"%s\"%s>", finalURL, targetAttr)
	}

	imgAttrs := fmt.Sprintf("src=\"%s\" alt=\"%s\" border=\"%s\"", selected.AdURL, selected.AltText, border)
	if width != "" {
		imgAttrs += fmt.Sprintf(" width=\"%s\"", width)
	}
	if height != "" {
		imgAttrs += fmt.Sprintf(" height=\"%s\"", height)
	}
	html += fmt.Sprintf("<img %s>", imgAttrs)

	if lib.Clickable && selected.HomeURL != "-" && selected.HomeURL != "" {
		html += "</a>"
	}

	return html
}

// --- MSWC.BrowserType ---

type G3BrowserType struct {
	vm         *VM
	properties map[string]interface{}
}

func (vm *VM) newG3BrowserTypeObject() Value {
	lib := &G3BrowserType{
		vm:         vm,
		properties: make(map[string]interface{}),
	}
	lib.detect()
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.mswcBrowserTypeItems[id] = lib
	return Value{Type: VTNativeObject, Num: id}
}

func (lib *G3BrowserType) detect() {
	ua := ""
	if lib.vm.host != nil && lib.vm.host.Request() != nil {
		ua = lib.vm.host.Request().ServerVars.Get("HTTP_USER_AGENT")
	}

	lib.properties["browser"] = "Unknown"
	lib.properties["version"] = "0.0"
	lib.properties["majorver"] = "0"
	lib.properties["minorver"] = "0"
	lib.properties["frames"] = false
	lib.properties["tables"] = true
	lib.properties["cookies"] = true
	lib.properties["backgroundsounds"] = false
	lib.properties["vbscript"] = false
	lib.properties["javascript"] = true
	lib.properties["javaapplets"] = false
	lib.properties["activexcontrols"] = false
	lib.properties["cdf"] = false

	if ua == "" {
		return
	}

	uaLower := strings.ToLower(ua)
	if strings.Contains(uaLower, "msie") || strings.Contains(uaLower, "trident") {
		lib.properties["browser"] = "IE"
		lib.properties["vbscript"] = true
		lib.properties["activexcontrols"] = true
		lib.properties["frames"] = true
		lib.properties["backgroundsounds"] = true
	} else if strings.Contains(uaLower, "edge") || strings.Contains(uaLower, "edg/") {
		lib.properties["browser"] = "Edge"
		lib.properties["frames"] = true
	} else if strings.Contains(uaLower, "chrome") {
		lib.properties["browser"] = "Chrome"
		lib.properties["frames"] = true
	} else if strings.Contains(uaLower, "firefox") {
		lib.properties["browser"] = "Firefox"
		lib.properties["frames"] = true
	} else if strings.Contains(uaLower, "safari") {
		lib.properties["browser"] = "Safari"
		lib.properties["frames"] = true
	}
}

func (lib *G3BrowserType) DispatchPropertyGet(name string) Value {
	name = strings.ToLower(name)
	if val, ok := lib.properties[name]; ok {
		switch v := val.(type) {
		case bool:
			return NewBool(v)
		case string:
			return NewString(v)
		}
	}
	return lib.DispatchMethod(name, nil)
}

func (lib *G3BrowserType) DispatchMethod(name string, args []Value) Value {
	return NewEmpty()
}

// --- MSWC.NextLink ---

type G3NextLink struct {
	vm *VM
}

func (vm *VM) newG3NextLinkObject() Value {
	obj := &G3NextLink{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.mswcNextLinkItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

func (lib *G3NextLink) DispatchPropertyGet(name string) Value {
	return lib.DispatchMethod(name, nil)
}

func (lib *G3NextLink) DispatchMethod(name string, args []Value) Value {
	if len(args) < 1 {
		return NewEmpty()
	}
	path := args[0].String()

	switch strings.ToLower(name) {
	case "getlistcount":
		return NewInteger(int64(lib.getListCount(path)))
	case "getlistindex":
		return NewInteger(int64(lib.getListIndex(path)))
	case "getnextdescription":
		return NewString(lib.getNextDescription(path))
	case "getnexturl":
		return NewString(lib.getNextURL(path))
	case "getpreviousdescription":
		return NewString(lib.getPreviousDescription(path))
	case "getpreviousurl":
		return NewString(lib.getPreviousURL(path))
	}

	if len(args) >= 2 {
		idx := int(lib.vm.asInt(args[1]))
		switch strings.ToLower(name) {
		case "getnthdescription":
			return NewString(lib.getNthDescription(path, idx))
		case "getnthurl":
			return NewString(lib.getNthURL(path, idx))
		}
	}

	return NewEmpty()
}

type linkEntry struct {
	URL         string
	Description string
	Comment     string
}

func (lib *G3NextLink) loadLinks(listFile string) []linkEntry {
	if lib.vm.host == nil || lib.vm.host.Server() == nil {
		return nil
	}
	absPath := lib.vm.host.Server().MapPath(listFile)
	file, err := os.Open(absPath)
	if err != nil {
		return nil
	}
	defer file.Close()

	var links []linkEntry
	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		line := scanner.Text()
		if strings.TrimSpace(line) == "" {
			continue
		}
		parts := strings.Split(line, "\t")
		entry := linkEntry{}
		if len(parts) > 0 {
			entry.URL = strings.TrimSpace(parts[0])
		}
		if len(parts) > 1 {
			entry.Description = strings.TrimSpace(parts[1])
		}
		if len(parts) > 2 {
			entry.Comment = strings.TrimSpace(parts[2])
		}
		if entry.URL != "" {
			links = append(links, entry)
		}
	}
	return links
}

func (lib *G3NextLink) getCurrentIndex(links []linkEntry) int {
	currentURL := ""
	if lib.vm.host != nil && lib.vm.host.Request() != nil {
		currentURL = lib.vm.host.Request().ServerVars.Get("SCRIPT_NAME")
	}
	for i, link := range links {
		if strings.HasSuffix(currentURL, link.URL) || strings.Contains(currentURL, link.URL) {
			return i + 1
		}
	}
	return 0
}

func (lib *G3NextLink) getListCount(path string) int {
	return len(lib.loadLinks(path))
}

func (lib *G3NextLink) getListIndex(path string) int {
	return lib.getCurrentIndex(lib.loadLinks(path))
}

func (lib *G3NextLink) getNextURL(path string) string {
	links := lib.loadLinks(path)
	idx := lib.getCurrentIndex(links)
	if idx > 0 && idx < len(links) {
		return links[idx].URL
	}
	return ""
}

func (lib *G3NextLink) getNextDescription(path string) string {
	links := lib.loadLinks(path)
	idx := lib.getCurrentIndex(links)
	if idx > 0 && idx < len(links) {
		return links[idx].Description
	}
	return ""
}

func (lib *G3NextLink) getPreviousURL(path string) string {
	links := lib.loadLinks(path)
	idx := lib.getCurrentIndex(links)
	if idx > 1 {
		return links[idx-2].URL
	}
	return ""
}

func (lib *G3NextLink) getPreviousDescription(path string) string {
	links := lib.loadLinks(path)
	idx := lib.getCurrentIndex(links)
	if idx > 1 {
		return links[idx-2].Description
	}
	return ""
}

func (lib *G3NextLink) getNthURL(path string, index int) string {
	links := lib.loadLinks(path)
	if index > 0 && index <= len(links) {
		return links[index-1].URL
	}
	return ""
}

func (lib *G3NextLink) getNthDescription(path string, index int) string {
	links := lib.loadLinks(path)
	if index > 0 && index <= len(links) {
		return links[index-1].Description
	}
	return ""
}

// --- MSWC.ContentRotator ---

type G3ContentRotator struct {
	vm *VM
}

func (vm *VM) newG3ContentRotatorObject() Value {
	obj := &G3ContentRotator{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.mswcContentRotatorItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

func (lib *G3ContentRotator) DispatchPropertyGet(name string) Value {
	return lib.DispatchMethod(name, nil)
}

func (lib *G3ContentRotator) DispatchMethod(name string, args []Value) Value {
	if len(args) < 1 {
		return NewEmpty()
	}
	path := args[0].String()

	switch strings.ToLower(name) {
	case "choosecontent":
		return NewString(lib.chooseContent(path))
	case "getallcontent":
		return NewString(lib.getAllContent(path))
	}
	return NewEmpty()
}

type contentEntry struct {
	Content string
	Weight  int
}

func (lib *G3ContentRotator) loadContent(path string) []contentEntry {
	if lib.vm.host == nil || lib.vm.host.Server() == nil {
		return nil
	}
	absPath := lib.vm.host.Server().MapPath(path)
	file, err := os.Open(absPath)
	if err != nil {
		return nil
	}
	defer file.Close()

	var entries []contentEntry
	var currentLines []string
	var currentWeight int = 1
	appendCurrentEntry := func() {
		content := strings.TrimSpace(strings.Join(currentLines, "\n"))
		if content == "" {
			return
		}
		entries = append(entries, contentEntry{
			Content: content,
			Weight:  currentWeight,
		})
	}

	scanner := bufio.NewScanner(file)
	firstEntry := true

	for scanner.Scan() {
		line := scanner.Text()
		trimmedLine := strings.TrimSpace(line)

		if strings.HasPrefix(trimmedLine, "%%") {
			if !firstEntry || len(currentLines) > 0 {
				appendCurrentEntry()
			}
			currentLines = nil
			currentWeight = 1
			firstEntry = false
			continue
		}

		if len(currentLines) == 0 && strings.HasPrefix(trimmedLine, "#") {
			headerParts := strings.SplitN(trimmedLine, "//", 2)
			weightPart := strings.TrimSpace(headerParts[0][1:])
			if w, err := strconv.Atoi(weightPart); err == nil {
				currentWeight = w
				continue
			}
		}

		currentLines = append(currentLines, line)
	}

	if len(currentLines) > 0 || !firstEntry {
		appendCurrentEntry()
	}

	return entries
}

func (lib *G3ContentRotator) chooseContent(path string) string {
	entries := lib.loadContent(path)
	if len(entries) == 0 {
		return ""
	}

	totalWeight := 0
	for _, e := range entries {
		if e.Weight > 0 {
			totalWeight += e.Weight
		}
	}

	if totalWeight <= 0 {
		r := rand.New(rand.NewSource(time.Now().UnixNano()))
		return entries[r.Intn(len(entries))].Content
	}

	r := rand.New(rand.NewSource(time.Now().UnixNano()))
	target := r.Intn(totalWeight)
	current := 0
	for _, e := range entries {
		if e.Weight <= 0 {
			continue
		}
		current += e.Weight
		if target < current {
			return e.Content
		}
	}
	return entries[0].Content
}

func (lib *G3ContentRotator) getAllContent(path string) string {
	entries := lib.loadContent(path)
	var sb strings.Builder
	for i, e := range entries {
		sb.WriteString(e.Content)
		if i < len(entries)-1 {
			sb.WriteString("<hr>\n")
		}
	}
	return sb.String()
}

// --- MSWC.Counters ---

var (
	countersMap  sync.Map
	countersFile string
	countersOnce sync.Once
)

type G3Counters struct {
	vm *VM
}

func (vm *VM) newG3CountersObject() Value {
	countersOnce.Do(func() {
		countersFile = filepath.Join("temp", "counters.txt")
		loadCounters()
	})
	obj := &G3Counters{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.mswcCountersItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

func loadCounters() {
	file, err := os.Open(countersFile)
	if err != nil {
		return
	}
	defer file.Close()

	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		line := scanner.Text()
		parts := strings.SplitN(line, "=", 2)
		if len(parts) == 2 {
			val, _ := strconv.Atoi(parts[1])
			countersMap.Store(strings.ToLower(parts[0]), val)
		}
	}
}

func saveCounters() {
	os.MkdirAll("temp", 0755)
	file, err := os.Create(countersFile)
	if err != nil {
		return
	}
	defer file.Close()

	countersMap.Range(func(key, value interface{}) bool {
		fmt.Fprintf(file, "%s=%v\n", key, value)
		return true
	})
}

func (lib *G3Counters) DispatchPropertyGet(name string) Value {
	return lib.DispatchMethod(name, nil)
}

func (lib *G3Counters) DispatchMethod(name string, args []Value) Value {
	if len(args) < 1 {
		return NewEmpty()
	}
	counterName := strings.ToLower(args[0].String())

	switch strings.ToLower(name) {
	case "get":
		val, ok := countersMap.Load(counterName)
		if ok {
			return NewInteger(int64(val.(int)))
		}
		return NewInteger(0)
	case "increment":
		val, _ := countersMap.LoadOrStore(counterName, 0)
		newVal := val.(int) + 1
		countersMap.Store(counterName, newVal)
		saveCounters()
		return NewInteger(int64(newVal))
	case "remove":
		countersMap.Delete(counterName)
		saveCounters()
		return NewEmpty()
	case "set":
		if len(args) >= 2 {
			newVal := int(lib.vm.asInt(args[1]))
			countersMap.Store(counterName, newVal)
			saveCounters()
		}
		return NewEmpty()
	}
	return NewEmpty()
}

// --- MSWC.Tools ---

type G3Tools struct {
	vm *VM
}

func (vm *VM) newG3ToolsObject() Value {
	obj := &G3Tools{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.mswcToolsItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

func (lib *G3Tools) DispatchPropertyGet(name string) Value {
	return lib.DispatchMethod(name, nil)
}

func (lib *G3Tools) DispatchMethod(name string, args []Value) Value {
	switch strings.ToLower(name) {
	case "fileexists":
		if len(args) < 1 || lib.vm.host == nil || lib.vm.host.Server() == nil {
			return NewBool(false)
		}
		path := lib.vm.host.Server().MapPath(args[0].String())
		_, err := os.Stat(path)
		return NewBool(err == nil)
	case "owner":
		if len(args) < 1 || lib.vm.host == nil || lib.vm.host.Server() == nil {
			return NewString("")
		}
		path := lib.vm.host.Server().MapPath(args[0].String())
		owner := lib.getFileOwner(path)
		return NewString(owner)
	case "pluginexists":
		return NewBool(false)
	case "processform":
		return NewEmpty()
	}
	return NewEmpty()
}

func (lib *G3Tools) getFileOwner(path string) string {
	ownerName := getFileOwnerName(path)
	if ownerName != "" {
		return ownerName
	}

	currentUser, err := user.Current()
	if err == nil {
		if currentUser.Name != "" {
			return currentUser.Name
		}
		if currentUser.Username != "" {
			parts := strings.Split(currentUser.Username, "\\")
			return parts[len(parts)-1]
		}
	}

	fallbacks := []string{"USERNAME", "USER", "LOGNAME"}
	for _, env := range fallbacks {
		if val := os.Getenv(env); val != "" {
			return val
		}
	}

	if runtime.GOOS == "windows" {
		return "System"
	}

	return ""
}

// --- MSWC.MyInfo ---

type G3MyInfo struct {
	vm         *VM
	properties map[string]string
	mu         sync.RWMutex
}

func (vm *VM) newG3MyInfoObject() Value {
	lib := &G3MyInfo{
		vm:         vm,
		properties: make(map[string]string),
	}
	lib.load()
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.mswcMyInfoItems[id] = lib
	return Value{Type: VTNativeObject, Num: id}
}

func (lib *G3MyInfo) load() {
	lib.mu.Lock()
	defer lib.mu.Unlock()

	if lib.vm.host == nil || lib.vm.host.Server() == nil {
		return
	}

	path := lib.vm.host.Server().MapPath("MyInfo.xml")
	data, err := os.ReadFile(path)
	if err != nil {
		return
	}

	type Property struct {
		XMLName xml.Name
		Value   string `xml:",chardata"`
	}
	type MyInfoXML struct {
		XMLName    xml.Name   `xml:"MyInfo"`
		Properties []Property `xml:",any"`
	}

	var config MyInfoXML
	if err := xml.Unmarshal(data, &config); err != nil {
		return
	}

	for _, p := range config.Properties {
		lib.properties[strings.ToLower(p.XMLName.Local)] = p.Value
	}
}

func (lib *G3MyInfo) DispatchPropertyGet(name string) Value {
	lib.mu.RLock()
	defer lib.mu.RUnlock()

	nameLower := strings.ToLower(name)
	if val, ok := lib.properties[nameLower]; ok {
		return NewString(val)
	}

	return lib.DispatchMethod(name, nil)
}

func (lib *G3MyInfo) DispatchPropertySet(name string, args []Value) bool {
	return false
}

func (lib *G3MyInfo) DispatchMethod(name string, args []Value) Value {
	lib.mu.RLock()
	defer lib.mu.RUnlock()

	nameLower := strings.ToLower(name)
	if nameLower == "url" || nameLower == "urlwords" {
		if len(args) > 0 {
			idx := args[0].String()
			propName := nameLower + idx
			if val, ok := lib.properties[propName]; ok {
				return NewString(val)
			}
		}
	} else if nameLower == "" && len(args) > 0 {
		propName := strings.ToLower(args[0].String())
		if val, ok := lib.properties[propName]; ok {
			return NewString(val)
		}
	}

	return NewEmpty()
}

// --- MSWC.PageCounter ---

var (
	pageCounterMap          sync.Map
	pageCounterFile         string
	pageCounterOnce         sync.Once
	pageCounterEnabled      bool
	pageCounterSaveInterval int
	pageCounterDirty        atomic.Bool
)

type G3PageCounter struct {
	vm *VM
}

func (vm *VM) newG3PageCounterObject() Value {
	pageCounterOnce.Do(func() {
		v := axonconfig.NewViper()

		// Preserve legacy behavior: MSWC page counter accepts env overrides regardless of global.viper_automatic_env.
		v.SetEnvKeyReplacer(strings.NewReplacer(".", "_"))
		v.AutomaticEnv()

		pageCounterEnabled = v.GetBool("mswc.pagecounter_enabled")
		pageCounterFile = v.GetString("mswc.pagecounter_file")
		if pageCounterFile == "" {
			pageCounterFile = filepath.Join("temp", "pagecounts.cnt")
		}
		pageCounterSaveInterval = v.GetInt("mswc.pagecounter_save_interval_seconds")
		if pageCounterSaveInterval <= 0 {
			pageCounterSaveInterval = 120
		}

		if pageCounterEnabled {
			loadPageCounts()
			startPageCounterBackgroundSaver()
		}
	})

	if !pageCounterEnabled {
		vm.raise(vbscript.InternalError, ErrPageCounterDisabled.String())
		return NewEmpty()
	}

	obj := &G3PageCounter{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.mswcPageCounterItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

func loadPageCounts() {
	file, err := os.Open(pageCounterFile)
	if err != nil {
		return
	}
	defer file.Close()

	var data map[string]int
	decoder := gob.NewDecoder(file)
	if err := decoder.Decode(&data); err == nil {
		for k, v := range data {
			pageCounterMap.Store(k, v)
		}
		// Reset dirty flag after successful load
		pageCounterDirty.Store(false)
	}
}

func savePageCounts() {
	// Optimization: check if data was modified before proceeding.
	if !pageCounterDirty.Swap(false) {
		return
	}

	// Fast Snapshotting: copy current key-value pairs to a temporary local map.
	snapshot := make(map[string]int)
	pageCounterMap.Range(func(key, value interface{}) bool {
		if k, ok := key.(string); ok {
			if v, ok := value.(int); ok {
				snapshot[k] = v
			}
		}
		return true
	})

	dir := filepath.Dir(pageCounterFile)
	os.MkdirAll(dir, 0755)
	file, err := os.Create(pageCounterFile)
	if err != nil {
		// If save fails, restore the dirty flag for the next attempt.
		pageCounterDirty.Store(true)
		return
	}
	defer file.Close()

	// Binary Storage using encoding/gob for maximum performance.
	encoder := gob.NewEncoder(file)
	_ = encoder.Encode(snapshot)
}

func startPageCounterBackgroundSaver() {
	go func() {
		ticker := time.NewTicker(time.Duration(pageCounterSaveInterval) * time.Second)
		defer ticker.Stop()
		for range ticker.C {
			savePageCounts()
		}
	}()
}

func (lib *G3PageCounter) DispatchPropertyGet(name string) Value {
	if !pageCounterEnabled {
		lib.vm.raise(vbscript.InternalError, ErrPageCounterDisabled.String())
		return NewEmpty()
	}
	return lib.DispatchMethod(name, nil)
}

func (lib *G3PageCounter) DispatchMethod(name string, args []Value) Value {
	if !pageCounterEnabled {
		lib.vm.raise(vbscript.InternalError, ErrPageCounterDisabled.String())
		return NewEmpty()
	}
	switch strings.ToLower(name) {
	case "hits":
		path := ""
		if len(args) >= 1 {
			path = args[0].String()
		} else if lib.vm.host != nil && lib.vm.host.Request() != nil {
			path = lib.vm.host.Request().ServerVars.Get("SCRIPT_NAME")
			if path == "" {
				path = lib.vm.host.Request().ServerVars.Get("URL")
			}
		}
		if path == "" {
			return NewInteger(0)
		}
		path = strings.ToLower(path)
		val, ok := pageCounterMap.Load(path)
		if ok {
			return NewInteger(int64(val.(int)))
		}
		return NewInteger(0)

	case "pagehit":
		path := ""
		if lib.vm.host != nil && lib.vm.host.Request() != nil {
			path = lib.vm.host.Request().ServerVars.Get("SCRIPT_NAME")
			if path == "" {
				path = lib.vm.host.Request().ServerVars.Get("URL")
			}
		}
		if path == "" {
			return NewInteger(0)
		}
		path = strings.ToLower(path)
		val, _ := pageCounterMap.LoadOrStore(path, 0)
		newVal := val.(int) + 1
		pageCounterMap.Store(path, newVal)
		// Mark data as modified to trigger disk write in background saver
		pageCounterDirty.Store(true)
		return NewInteger(int64(newVal))

	case "reset":
		path := ""
		if len(args) >= 1 {
			path = args[0].String()
		} else if lib.vm.host != nil && lib.vm.host.Request() != nil {
			path = lib.vm.host.Request().ServerVars.Get("SCRIPT_NAME")
			if path == "" {
				path = lib.vm.host.Request().ServerVars.Get("URL")
			}
		}
		if path != "" {
			pageCounterMap.Delete(strings.ToLower(path))
			// Mark data as modified to trigger disk write in background saver
			pageCounterDirty.Store(true)
		}
		return NewEmpty()
	}
	return NewEmpty()
}

// --- MSWC.PermissionChecker ---

// G3PermissionChecker implements the MSWC.PermissionChecker component.
// It allows scripts to verify if the current process has read access to a specific path.
type G3PermissionChecker struct {
	vm *VM
}

// newG3PermissionCheckerObject creates a new G3PermissionChecker instance.
func (vm *VM) newG3PermissionCheckerObject() Value {
	obj := &G3PermissionChecker{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.mswcPermissionCheckerItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet handles property read access for the component.
func (lib *G3PermissionChecker) DispatchPropertyGet(name string) Value {
	return lib.DispatchMethod(name, nil)
}

// DispatchMethod handles method calls for the component.
func (lib *G3PermissionChecker) DispatchMethod(name string, args []Value) Value {
	if len(args) < 1 {
		return NewEmpty()
	}
	resourcePath := args[0].String()

	switch strings.ToLower(name) {
	case "hasaccess":
		// HasAccess verifies if the script has read access to the specified path.
		return NewBool(lib.hasAccess(resourcePath))
	}
	return NewEmpty()
}

// hasAccess evaluates read access for a virtual or physical path.
func (lib *G3PermissionChecker) hasAccess(resourcePath string) bool {
	if lib.vm.host == nil || lib.vm.host.Server() == nil {
		return false
	}

	// Map the path to a physical location.
	absPath := lib.vm.host.Server().MapPath(resourcePath)

	// Attempt a read-only open to verify actual read access.
	// This approach is cross-platform (Windows, macOS, Linux) and avoids complex OS-specific branches.
	// It directly tests if the OS allows the running process to read the target file or directory.
	f, err := os.Open(absPath)
	if err != nil {
		// If we can't open it (e.g., Access Denied, Not Found), return false.
		return false
	}
	f.Close()

	return true
}
