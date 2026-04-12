/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimaraes - G3pix Ltda
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
	"sort"
	"strings"

	"g3pix.com.br/axonasp/axonvm/asp"
	"g3pix.com.br/axonasp/vbscript"
)

// G3TestFailure stores one failed assertion message produced by one suite run.
type G3TestFailure struct {
	Message string
}

// G3TestReport is the immutable runner-facing snapshot of one G3Test object.
type G3TestReport struct {
	SuiteName string
	Total     int64
	Passed    int64
	Failed    int64
	Failures  []G3TestFailure
}

// G3TestSuiteSummary aggregates all G3Test objects created in one VM execution.
type G3TestSuiteSummary struct {
	SuiteCount int
	Total      int64
	Passed     int64
	Failed     int64
	Failures   []G3TestFailure
}

// G3Test is the native assertion object used by axonasp-testsuite ASP tests.
type G3Test struct {
	vm           *VM
	currentSuite string
	total        int64
	passed       int64
	failed       int64
	failures     []G3TestFailure
}

// newG3TestObject allocates and registers one G3Test object in the VM native-object table.
func (vm *VM) newG3TestObject() Value {
	obj := &G3Test{vm: vm, failures: make([]G3TestFailure, 0, 8)}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3testItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchMethod routes method calls with a case-insensitive O(1) switch.
func (g *G3Test) DispatchMethod(methodName string, args []Value) Value {
	switch strings.ToLower(strings.TrimSpace(methodName)) {
	case "assertequal":
		return g.methodAssertEqual(args)
	case "assertnotequal":
		return g.methodAssertNotEqual(args)
	case "asserttrue":
		return g.methodAssertTrue(args)
	case "assertfalse":
		return g.methodAssertFalse(args)
	case "assertempty":
		return g.methodAssertEmpty(args)
	case "assertnull":
		return g.methodAssertNull(args)
	case "assertnothing":
		return g.methodAssertNothing(args)
	case "asserttypename":
		return g.methodAssertTypeName(args)
	case "assertlength", "assertcount":
		return g.methodAssertLength(args)
	case "assertraises":
		return g.methodAssertRaises(args)
	case "fail":
		return g.methodFail(args)
	case "describe":
		return g.methodDescribe(args)
	}
	return NewEmpty()
}

// DispatchPropertyGet returns immutable suite counters and metadata.
func (g *G3Test) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(strings.TrimSpace(propertyName)) {
	case "suite", "description", "currentdescribe":
		return NewString(g.currentSuite)
	case "total", "totaltests":
		return NewInteger(g.total)
	case "passed":
		return NewInteger(g.passed)
	case "failed":
		return NewInteger(g.failed)
	case "hasfailures":
		return NewBool(g.failed > 0)
	}
	return g.DispatchMethod(propertyName, nil)
}

// DispatchPropertySet updates writable suite metadata fields.
func (g *G3Test) DispatchPropertySet(propertyName string, val Value) {
	switch strings.ToLower(strings.TrimSpace(propertyName)) {
	case "suite", "description", "currentdescribe":
		g.currentSuite = strings.TrimSpace(val.String())
	}
}

// Snapshot returns one copy-only report that is safe to consume after VM execution.
func (g *G3Test) Snapshot() G3TestReport {
	report := G3TestReport{
		SuiteName: g.currentSuite,
		Total:     g.total,
		Passed:    g.passed,
		Failed:    g.failed,
		Failures:  make([]G3TestFailure, len(g.failures)),
	}
	copy(report.Failures, g.failures)
	return report
}

// GetG3TestReports returns all test object reports in stable object-id order.
func (vm *VM) GetG3TestReports() []G3TestReport {
	if vm == nil || len(vm.g3testItems) == 0 {
		return nil
	}
	ids := make([]int64, 0, len(vm.g3testItems))
	for id := range vm.g3testItems {
		ids = append(ids, id)
	}
	sort.Slice(ids, func(i int, j int) bool { return ids[i] < ids[j] })

	reports := make([]G3TestReport, 0, len(ids))
	for _, id := range ids {
		obj := vm.g3testItems[id]
		if obj == nil {
			continue
		}
		reports = append(reports, obj.Snapshot())
	}
	return reports
}

// GetG3TestSuiteSummary merges all G3Test object reports from the current VM.
func (vm *VM) GetG3TestSuiteSummary() G3TestSuiteSummary {
	reports := vm.GetG3TestReports()
	if len(reports) == 0 {
		return G3TestSuiteSummary{}
	}
	failureCap := 0
	for i := range reports {
		failureCap += len(reports[i].Failures)
	}
	summary := G3TestSuiteSummary{SuiteCount: len(reports), Failures: make([]G3TestFailure, 0, failureCap)}
	for i := range reports {
		summary.Total += reports[i].Total
		summary.Passed += reports[i].Passed
		summary.Failed += reports[i].Failed
		summary.Failures = append(summary.Failures, reports[i].Failures...)
	}
	return summary
}

// methodDescribe sets the active test block/suite name for assertion diagnostics.
func (g *G3Test) methodDescribe(args []Value) Value {
	if len(args) == 0 {
		g.currentSuite = ""
		return NewEmpty()
	}
	g.currentSuite = strings.TrimSpace(args[0].String())
	return NewEmpty()
}

// methodAssertEqual validates equality using VM comparison/coercion semantics.
func (g *G3Test) methodAssertEqual(args []Value) Value {
	if len(args) < 2 {
		return g.recordFailure("AssertEqual requires expected and actual values")
	}
	if g.vm == nil {
		return g.recordFailure("AssertEqual requires an active VM")
	}
	expected := resolveCallable(g.vm, args[0])
	actual := resolveCallable(g.vm, args[1])
	msg := ""
	if len(args) > 2 {
		msg = strings.TrimSpace(args[2].String())
	}

	g.total++
	if g.valuesEqual(expected, actual) {
		g.passed++
		return NewBool(true)
	}

	detail := "AssertEqual failed"
	if msg != "" {
		detail = msg
	}
	detail += ": expected=" + g.valuePreview(expected) + " actual=" + g.valuePreview(actual)
	g.failed++
	g.appendFailure(detail)
	return NewBool(false)
}

// methodAssertNotEqual validates inequality using VM comparison/coercion semantics.
func (g *G3Test) methodAssertNotEqual(args []Value) Value {
	if len(args) < 2 {
		return g.recordFailure("AssertNotEqual requires expected and actual values")
	}
	if g.vm == nil {
		return g.recordFailure("AssertNotEqual requires an active VM")
	}
	expected := resolveCallable(g.vm, args[0])
	actual := resolveCallable(g.vm, args[1])
	msg := "AssertNotEqual failed"
	if len(args) > 2 {
		if text := strings.TrimSpace(args[2].String()); text != "" {
			msg = text
		}
	}

	g.total++
	if !g.valuesEqual(expected, actual) {
		g.passed++
		return NewBool(true)
	}

	g.failed++
	g.appendFailure(msg + ": both=" + g.valuePreview(actual))
	return NewBool(false)
}

// methodAssertTrue validates one truthy condition using VM boolean coercion.
func (g *G3Test) methodAssertTrue(args []Value) Value {
	if len(args) == 0 {
		return g.recordFailure("AssertTrue requires a condition value")
	}
	if g.vm == nil {
		return g.recordFailure("AssertTrue requires an active VM")
	}
	g.total++
	if g.vm.asBool(resolveCallable(g.vm, args[0])) {
		g.passed++
		return NewBool(true)
	}
	msg := "AssertTrue failed"
	if len(args) > 1 {
		if text := strings.TrimSpace(args[1].String()); text != "" {
			msg = text
		}
	}
	g.failed++
	g.appendFailure(msg)
	return NewBool(false)
}

// methodAssertFalse validates one falsy condition using VM boolean coercion.
func (g *G3Test) methodAssertFalse(args []Value) Value {
	if len(args) == 0 {
		return g.recordFailure("AssertFalse requires a condition value")
	}
	if g.vm == nil {
		return g.recordFailure("AssertFalse requires an active VM")
	}
	g.total++
	if !g.vm.asBool(resolveCallable(g.vm, args[0])) {
		g.passed++
		return NewBool(true)
	}
	msg := "AssertFalse failed"
	if len(args) > 1 {
		if text := strings.TrimSpace(args[1].String()); text != "" {
			msg = text
		}
	}
	g.failed++
	g.appendFailure(msg)
	return NewBool(false)
}

// methodAssertEmpty validates that one value is the VBScript Empty variant.
func (g *G3Test) methodAssertEmpty(args []Value) Value {
	if len(args) == 0 {
		return g.recordFailure("AssertEmpty requires a value")
	}
	if g.vm == nil {
		return g.recordFailure("AssertEmpty requires an active VM")
	}
	value := resolveCallable(g.vm, args[0])
	msg := "AssertEmpty failed"
	if len(args) > 1 {
		if text := strings.TrimSpace(args[1].String()); text != "" {
			msg = text
		}
	}

	g.total++
	if isEmpty(value) {
		g.passed++
		return NewBool(true)
	}

	g.failed++
	g.appendFailure(msg + ": actual=" + g.valuePreview(value))
	return NewBool(false)
}

// methodAssertNull validates that one value is the VBScript Null variant.
func (g *G3Test) methodAssertNull(args []Value) Value {
	if len(args) == 0 {
		return g.recordFailure("AssertNull requires a value")
	}
	if g.vm == nil {
		return g.recordFailure("AssertNull requires an active VM")
	}
	value := resolveCallable(g.vm, args[0])
	msg := "AssertNull failed"
	if len(args) > 1 {
		if text := strings.TrimSpace(args[1].String()); text != "" {
			msg = text
		}
	}

	g.total++
	if isNull(value) {
		g.passed++
		return NewBool(true)
	}

	g.failed++
	g.appendFailure(msg + ": actual=" + g.valuePreview(value))
	return NewBool(false)
}

// methodAssertNothing validates that one value is Nothing-compatible in VBScript terms.
func (g *G3Test) methodAssertNothing(args []Value) Value {
	if len(args) == 0 {
		return g.recordFailure("AssertNothing requires a value")
	}
	if g.vm == nil {
		return g.recordFailure("AssertNothing requires an active VM")
	}
	value := resolveCallable(g.vm, args[0])
	msg := "AssertNothing failed"
	if len(args) > 1 {
		if text := strings.TrimSpace(args[1].String()); text != "" {
			msg = text
		}
	}

	g.total++
	if g.isNothingLike(value) {
		g.passed++
		return NewBool(true)
	}

	g.failed++
	g.appendFailure(msg + ": actual=" + g.valuePreview(value))
	return NewBool(false)
}

// methodAssertTypeName validates VBScript TypeName output for one value.
func (g *G3Test) methodAssertTypeName(args []Value) Value {
	if len(args) < 2 {
		return g.recordFailure("AssertTypeName requires expectedType and value")
	}
	if g.vm == nil {
		return g.recordFailure("AssertTypeName requires an active VM")
	}

	expectedType := strings.TrimSpace(g.vm.valueToString(resolveCallable(g.vm, args[0])))
	value := resolveCallable(g.vm, args[1])
	msg := "AssertTypeName failed"
	if len(args) > 2 {
		if text := strings.TrimSpace(args[2].String()); text != "" {
			msg = text
		}
	}

	g.total++
	actualType := g.resolveTypeName(value)
	if strings.EqualFold(expectedType, actualType) {
		g.passed++
		return NewBool(true)
	}

	g.failed++
	g.appendFailure(msg + ": expectedType=" + expectedType + " actualType=" + actualType)
	return NewBool(false)
}

// methodAssertLength validates one length/count value from arrays, strings, and collections.
func (g *G3Test) methodAssertLength(args []Value) Value {
	if len(args) < 2 {
		return g.recordFailure("AssertLength requires expectedLength and value")
	}
	if g.vm == nil {
		return g.recordFailure("AssertLength requires an active VM")
	}

	expected := g.vm.asInt(resolveCallable(g.vm, args[0]))
	value := resolveCallable(g.vm, args[1])
	msg := "AssertLength failed"
	if len(args) > 2 {
		if text := strings.TrimSpace(args[2].String()); text != "" {
			msg = text
		}
	}

	g.total++
	actual, ok := g.resolveLength(value)
	if !ok {
		g.failed++
		g.appendFailure(msg + ": value does not expose length/count")
		return NewBool(false)
	}
	if actual == expected {
		g.passed++
		return NewBool(true)
	}

	g.failed++
	g.appendFailure(msg + ": expectedLength=" + g.intToString(expected) + " actualLength=" + g.intToString(actual))
	return NewBool(false)
}

// resolveTypeName returns the VBScript TypeName output for one value.
func (g *G3Test) resolveTypeName(value Value) string {
	if g.vm == nil {
		return "Unknown"
	}
	result, err := vbsTypeNameVM(g.vm, []Value{value})
	if err != nil {
		return "Unknown"
	}
	return g.vm.valueToString(result)
}

// resolveLength extracts one count value from strings, arrays, and collection-like native objects.
func (g *G3Test) resolveLength(value Value) (int, bool) {
	if g.vm == nil {
		return 0, false
	}

	switch value.Type {
	case VTString:
		return len([]rune(g.vm.valueToString(value))), true
	case VTArray:
		if value.Arr == nil {
			return 0, true
		}
		return len(value.Arr.Values), true
	case VTEmpty:
		return 0, true
	case VTNativeObject, VTObject:
		countValue := g.vm.dispatchMemberGet(value, "Count")
		if isNumericLike(countValue) || countValue.Type == VTString {
			return g.vm.asInt(countValue), true
		}
		return 0, false
	default:
		return 0, false
	}
}

// isNothingLike returns true when one value matches the project-compatible Nothing semantics.
func (g *G3Test) isNothingLike(value Value) bool {
	if value.Type == VTObject && value.Num == 0 {
		return true
	}
	if value.Type == VTEmpty || value.Type == VTNull {
		return true
	}
	return false
}

// methodAssertRaises validates that one Execute-compatible VBScript code block raises an error.
func (g *G3Test) methodAssertRaises(args []Value) Value {
	if len(args) == 0 {
		return g.recordFailure("AssertRaises requires VBScript code")
	}
	if g.vm == nil {
		return g.recordFailure("AssertRaises requires an active VM")
	}

	code := strings.TrimSpace(g.vm.valueToString(resolveCallable(g.vm, args[0])))
	if code == "" {
		return g.recordFailure("AssertRaises requires non-empty VBScript code")
	}

	expectedNumber, expectedText, message := g.parseAssertRaisesOptions(args)
	g.total++

	raised := g.captureDynamicError(code)
	if raised == nil {
		g.failed++
		g.appendFailure(message + ": no error was raised")
		return NewBool(false)
	}
	if expectedNumber != 0 && raised.Number != expectedNumber {
		g.failed++
		g.appendFailure(message + ": expected Err.Number=" + g.intToString(expectedNumber) + " actual=" + g.intToString(raised.Number))
		return NewBool(false)
	}
	if expectedText != "" && !g.vmErrorContains(raised, expectedText) {
		g.failed++
		g.appendFailure(message + ": expected error text containing=" + expectedText + " actual=" + g.vmErrorPreview(raised))
		return NewBool(false)
	}

	g.passed++
	return NewBool(true)
}

// parseAssertRaisesOptions resolves the optional expected error and message arguments.
func (g *G3Test) parseAssertRaisesOptions(args []Value) (int, string, string) {
	expectedNumber := 0
	expectedText := ""
	message := "AssertRaises failed"

	if g.vm == nil || len(args) < 2 {
		return expectedNumber, expectedText, message
	}

	if len(args) == 2 {
		candidate := resolveCallable(g.vm, args[1])
		if isNumericLike(candidate) {
			expectedNumber = g.vm.asInt(candidate)
			return expectedNumber, expectedText, message
		}
		if text := strings.TrimSpace(g.vm.valueToString(candidate)); text != "" {
			message = text
		}
		return expectedNumber, expectedText, message
	}

	expected := resolveCallable(g.vm, args[1])
	if isNumericLike(expected) {
		expectedNumber = g.vm.asInt(expected)
	} else {
		expectedText = strings.TrimSpace(g.vm.valueToString(expected))
	}
	if text := strings.TrimSpace(g.vm.valueToString(resolveCallable(g.vm, args[2]))); text != "" {
		message = text
	}
	return expectedNumber, expectedText, message
}

// captureDynamicError executes one VBScript statement block and captures a raised VM error without leaking Err state.
func (g *G3Test) captureDynamicError(code string) *VMError {
	if g.vm == nil {
		return nil
	}

	previousErr := asp.NewASPError()
	if g.vm.errObject != nil {
		previousErr = g.vm.errObject.Clone()
	}
	previousRaw := g.vm.errASPCodeRaw
	previousRawSet := g.vm.errASPCodeRawSet
	previousLastError := g.vm.lastError

	g.vm.errClear()
	defer func() {
		g.vm.errObject = previousErr
		g.vm.errASPCodeRaw = previousRaw
		g.vm.errASPCodeRawSet = previousRawSet
		g.vm.lastError = previousLastError
	}()

	var raised *VMError
	func() {
		defer func() {
			recovered := recover()
			if recovered == nil {
				return
			}
			switch typed := recovered.(type) {
			case *VMError:
				raised = typed
			case error:
				if vme, ok := typed.(*VMError); ok {
					raised = vme
					return
				}
				panic(recovered)
			default:
				panic(recovered)
			}
		}()

		_, err := vbsCompatExecute(g.vm, []Value{NewString(code)})
		if err == nil {
			return
		}
		if vme, ok := err.(*VMError); ok {
			raised = vme
			return
		}
		raised = g.newInternalAssertionVMError(err.Error())
	}()

	if raised != nil {
		return raised
	}
	if g.vm.errObject != nil && g.vm.errObject.Number != 0 {
		return &VMError{
			ASPCode:        g.vm.errObject.ASPCode,
			ASPDescription: g.vm.errObject.ASPDescription,
			Number:         g.vm.errObject.Number,
			Source:         g.vm.errObject.Source,
			Description:    g.vm.errObject.Description,
			HelpFile:       g.vm.errObject.HelpFile,
			HelpContext:    g.vm.errObject.HelpContext,
			File:           g.vm.errObject.File,
			Line:           g.vm.errObject.Line,
			Column:         g.vm.errObject.Column,
			Category:       g.vm.errObject.Category,
		}
	}
	return nil
}

// newInternalAssertionVMError converts one unexpected Go error into a VMError-compatible payload.
func (g *G3Test) newInternalAssertionVMError(message string) *VMError {
	line := 0
	column := 0
	if g.vm != nil {
		line = g.vm.lastLine
		column = g.vm.lastColumn
	}
	return &VMError{
		Code:           vbscript.InternalError,
		Line:           line,
		Column:         column,
		Msg:            message,
		ASPCode:        int(vbscript.InternalError),
		ASPDescription: message,
		Category:       "VBScript runtime",
		Description:    message,
		Number:         vbscript.HRESULTFromVBScriptCode(vbscript.InternalError),
		Source:         "VBScript runtime error",
	}
}

// vmErrorContains reports whether one VM error contains the expected text in core fields.
func (g *G3Test) vmErrorContains(vme *VMError, expected string) bool {
	if vme == nil {
		return false
	}
	needle := strings.ToLower(strings.TrimSpace(expected))
	if needle == "" {
		return true
	}
	haystack := strings.ToLower(vme.Description + "\n" + vme.ASPDescription + "\n" + vme.Source + "\n" + vme.Msg)
	return strings.Contains(haystack, needle)
}

// vmErrorPreview creates one compact description of a trapped VM error for failure output.
func (g *G3Test) vmErrorPreview(vme *VMError) string {
	if vme == nil {
		return "<nil>"
	}
	if text := strings.TrimSpace(vme.Description); text != "" {
		return text
	}
	if text := strings.TrimSpace(vme.ASPDescription); text != "" {
		return text
	}
	if text := strings.TrimSpace(vme.Msg); text != "" {
		return text
	}
	return vme.Source
}

// intToString formats one integer for assertion diagnostics without extra allocations outside fmt.
func (g *G3Test) intToString(value int) string {
	return NewInteger(int64(value)).String()
}

// methodFail marks one explicit assertion failure.
func (g *G3Test) methodFail(args []Value) Value {
	g.total++
	g.failed++
	msg := "Fail called"
	if len(args) > 0 {
		if text := strings.TrimSpace(args[0].String()); text != "" {
			msg = text
		}
	}
	g.appendFailure(msg)
	return NewBool(false)
}

// recordFailure writes one failure for argument/usage errors and returns False.
func (g *G3Test) recordFailure(message string) Value {
	g.total++
	g.failed++
	g.appendFailure(message)
	return NewBool(false)
}

// appendFailure appends one formatted failure message with optional suite context.
func (g *G3Test) appendFailure(message string) {
	if g.currentSuite != "" {
		message = "[" + g.currentSuite + "] " + message
	}
	g.failures = append(g.failures, G3TestFailure{Message: message})
}

// valuePreview creates one short value representation for assertion diff output.
func (g *G3Test) valuePreview(v Value) string {
	if g.vm == nil {
		return v.String()
	}
	text := g.vm.valueToString(v)
	if text == "" {
		return "<empty>"
	}
	return text
}

// valuesEqual checks equality with the same core coercion strategy used by OpEq.
func (g *G3Test) valuesEqual(a Value, b Value) bool {
	if isNull(a) || isNull(b) {
		return false
	}
	if g.vm != nil && g.vm.optionCompare == 1 {
		return strings.EqualFold(a.String(), b.String())
	}
	if g.vm == nil {
		return a.String() == b.String()
	}
	return g.vm.compareValues(a, b) == 0
}
