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
	"regexp"
	"strings"
)

// ─── METADATA Directive Scanner ───────────────────────────────────────────────

// metadataDirectivePattern matches <!-- METADATA TYPE="TypeLib" ... --> HTML comments,
// capturing the full attribute block so the parser can extract UUID, NAME, and VERSION.
var metadataDirectivePattern = regexp.MustCompile(`(?is)<!--\s*METADATA[^>]*TYPE\s*=\s*["']?\s*TypeLib\s*["']?[^>]*-->`)

// metadataAttrPattern extracts a single attribute value (UUID, NAME, or VERSION) from
// a METADATA comment.
var metadataAttrPattern = regexp.MustCompile(`(?i)(?:UUID|NAME|VERSION)\s*=\s*["']([^"']*)["']`)

// stripMetadataDirectives removes all <!-- METADATA ... --> HTML comments from the
// source so they do not appear in the rendered page output, matching IIS behaviour.
func stripMetadataDirectives(source string) string {
	return metadataDirectivePattern.ReplaceAllLiteralString(source, "")
}

// detectedMetadataLibrary carries one parsed METADATA TYPE="TypeLib" entry.
type detectedMetadataLibrary struct {
	uuid string
	name string
}

// detectMetadataLibraries scans one ASP source and extracts all METADATA TYPE="TypeLib" entries.
// UUID is normalized to uppercase without braces and NAME is trimmed.
func detectMetadataLibraries(source string) []detectedMetadataLibrary {
	matches := metadataDirectivePattern.FindAllString(source, -1)
	if len(matches) == 0 {
		return nil
	}

	libs := make([]detectedMetadataLibrary, 0, len(matches))
	seen := make(map[string]bool)

	for _, block := range matches {
		lib := detectedMetadataLibrary{}
		for _, attr := range metadataAttrPattern.FindAllStringSubmatch(block, -1) {
			if len(attr) < 2 {
				continue
			}
			raw := strings.TrimSpace(attr[0])
			val := strings.TrimSpace(attr[1])
			upperRaw := strings.ToUpper(raw)
			switch {
			case strings.HasPrefix(upperRaw, "UUID"):
				lib.uuid = strings.ToUpper(strings.Trim(val, "{}"))
			case strings.HasPrefix(upperRaw, "NAME"):
				lib.name = val
			}
		}

		key := lib.uuid + "|" + strings.ToLower(lib.name)
		if lib.uuid == "" && strings.TrimSpace(lib.name) == "" {
			continue
		}
		if seen[key] {
			continue
		}
		seen[key] = true
		libs = append(libs, lib)
	}

	if len(libs) == 0 {
		return nil
	}
	return libs
}

// getMetadataLibraryConstants resolves one deduplicated constant set for detected typelibs.
// Matching uses UUID first, then common library-name aliases used by Classic ASP.
func getMetadataLibraryConstants(libs []detectedMetadataLibrary) []VBSConstant {
	if len(libs) == 0 {
		return nil
	}

	out := make([]VBSConstant, 0)
	seenConst := make(map[string]bool)

	appendSet := func(set []VBSConstant) {
		for _, kv := range set {
			key := strings.ToLower(kv.Name)
			if seenConst[key] {
				continue
			}
			seenConst[key] = true
			out = append(out, kv)
		}
	}

	for _, lib := range libs {
		uuid := strings.ToUpper(strings.TrimSpace(lib.uuid))
		name := strings.ToLower(strings.TrimSpace(lib.name))

		switch uuid {
		case "00000205-0000-0010-8000-00AA006D2EA4", // ADODB 2.5
			"00000200-0000-0010-8000-00AA006D2EA4", // ADODB 2.0
			"00000201-0000-0010-8000-00AA006D2EA4", // ADODB 2.1
			"00000202-0000-0010-8000-00AA006D2EA4", // ADODB 2.5 alias
			"00000513-0000-0010-8000-00AA006D2EA4": // ADOX
			appendSet(adodbTypeLibConstants)
			continue

		case "420B2830-E718-11CF-893D-00A0C9054228": // Microsoft Scripting Runtime (scrrun)
			appendSet(fsoTypeLibConstants)
			continue

		case "CD000000-8B95-11D1-82DB-00C04FB1625D": // CDO for Windows 2000
			appendSet(cdoTypeLibConstants)
			continue
		}

		if strings.Contains(name, "activex data objects") || strings.Contains(name, "adodb") || strings.Contains(name, "ado") {
			appendSet(adodbTypeLibConstants)
			continue
		}
		if strings.Contains(name, "scripting runtime") || strings.Contains(name, "microsoft scripting") || strings.Contains(name, "scrrun") {
			appendSet(fsoTypeLibConstants)
			continue
		}
		if strings.Contains(name, "collaboration data objects") || strings.Contains(name, "cdosys") || strings.Contains(name, "cdonts") || name == "cdo" {
			appendSet(cdoTypeLibConstants)
		}
	}

	if len(out) == 0 {
		return nil
	}
	return out
}

// ─── ADODB 2.5 Type Library Constants ─────────────────────────────────────────
// Source: Microsoft ADOVBS.INC / ADODB 2.5 Type Library (adovbs.inc, ADO 2.5 SDK).
// All constant names and values must match the official Microsoft documentation exactly.
var adodbTypeLibConstants = []VBSConstant{

	// ── DataTypeEnum: field/parameter data types ────────────────────────────
	{"adEmpty", NewInteger(0)},
	{"adTinyInt", NewInteger(16)},
	{"adSmallInt", NewInteger(2)},
	{"adInteger", NewInteger(3)},
	{"adBigInt", NewInteger(20)},
	{"adUnsignedTinyInt", NewInteger(17)},
	{"adUnsignedSmallInt", NewInteger(18)},
	{"adUnsignedInt", NewInteger(19)},
	{"adUnsignedBigInt", NewInteger(21)},
	{"adSingle", NewInteger(4)},
	{"adDouble", NewInteger(5)},
	{"adCurrency", NewInteger(6)},
	{"adDecimal", NewInteger(14)},
	{"adNumeric", NewInteger(131)},
	{"adBoolean", NewInteger(11)},
	{"adError", NewInteger(10)},
	{"adUserDefined", NewInteger(132)},
	{"adVariant", NewInteger(12)},
	{"adIDispatch", NewInteger(9)},
	{"adIUnknown", NewInteger(13)},
	{"adGUID", NewInteger(72)},
	{"adDate", NewInteger(7)},
	{"adDBDate", NewInteger(133)},
	{"adDBTime", NewInteger(134)},
	{"adDBTimeStamp", NewInteger(135)},
	{"adBSTR", NewInteger(8)},
	{"adChar", NewInteger(129)},
	{"adVarChar", NewInteger(200)},
	{"adLongVarChar", NewInteger(201)},
	{"adWChar", NewInteger(130)},
	{"adVarWChar", NewInteger(202)},
	{"adLongVarWChar", NewInteger(203)},
	{"adBinary", NewInteger(128)},
	{"adVarBinary", NewInteger(204)},
	{"adLongVarBinary", NewInteger(205)},
	{"adChapter", NewInteger(136)},
	{"adFileTime", NewInteger(64)},
	{"adPropVariant", NewInteger(138)},
	{"adVarNumeric", NewInteger(139)},
	{"adArray", NewInteger(8192)},

	// ── FieldAttributeEnum: field attribute flags ───────────────────────────
	{"adFldMayDefer", NewInteger(2)},
	{"adFldUpdatable", NewInteger(4)},
	{"adFldUnknownUpdatable", NewInteger(8)},
	{"adFldFixed", NewInteger(16)},
	{"adFldIsNullable", NewInteger(32)},
	{"adFldMayBeNull", NewInteger(64)},
	{"adFldLong", NewInteger(128)},
	{"adFldRowID", NewInteger(256)},
	{"adFldRowVersion", NewInteger(512)},
	{"adFldCacheDeferred", NewInteger(1024)},
	{"adFldNegativeScale", NewInteger(16384)},
	{"adFldKeyColumn", NewInteger(32768)},

	// ── CursorTypeEnum ──────────────────────────────────────────────────────
	{"adOpenForwardOnly", NewInteger(0)},
	{"adOpenKeyset", NewInteger(1)},
	{"adOpenDynamic", NewInteger(2)},
	{"adOpenStatic", NewInteger(3)},

	// ── LockTypeEnum ────────────────────────────────────────────────────────
	{"adLockUnspecified", NewInteger(-1)},
	{"adLockReadOnly", NewInteger(1)},
	{"adLockPessimistic", NewInteger(2)},
	{"adLockOptimistic", NewInteger(3)},
	{"adLockBatchOptimistic", NewInteger(4)},

	// ── CursorLocationEnum ──────────────────────────────────────────────────
	{"adUseNone", NewInteger(1)},
	{"adUseServer", NewInteger(2)},
	{"adUseClient", NewInteger(3)},
	{"adUseClientBatch", NewInteger(3)}, // Alias for adUseClient

	// ── ConnectModeEnum ─────────────────────────────────────────────────────
	{"adModeUnknown", NewInteger(0)},
	{"adModeRead", NewInteger(1)},
	{"adModeWrite", NewInteger(2)},
	{"adModeReadWrite", NewInteger(3)},
	{"adModeShareDenyRead", NewInteger(4)},
	{"adModeShareDenyWrite", NewInteger(8)},
	{"adModeShareExclusive", NewInteger(12)},
	{"adModeShareDenyNone", NewInteger(16)},
	{"adModeRecursive", NewInteger(4194304)},

	// ── ObjectStateEnum ─────────────────────────────────────────────────────
	{"adStateClosed", NewInteger(0)},
	{"adStateOpen", NewInteger(1)},
	{"adStateConnecting", NewInteger(2)},
	{"adStateExecuting", NewInteger(4)},
	{"adStateFetching", NewInteger(8)},

	// ── CommandTypeEnum ─────────────────────────────────────────────────────
	{"adCmdUnspecified", NewInteger(-1)},
	{"adCmdText", NewInteger(1)},
	{"adCmdTable", NewInteger(2)},
	{"adCmdStoredProc", NewInteger(4)},
	{"adCmdUnknown", NewInteger(8)},
	{"adCmdFile", NewInteger(256)},
	{"adCmdTableDirect", NewInteger(512)},

	// ── ParameterDirectionEnum ──────────────────────────────────────────────
	{"adParamUnknown", NewInteger(0)},
	{"adParamInput", NewInteger(1)},
	{"adParamOutput", NewInteger(2)},
	{"adParamInputOutput", NewInteger(3)},
	{"adParamReturnValue", NewInteger(4)},

	// ── FilterGroupEnum ─────────────────────────────────────────────────────
	{"adFilterNone", NewInteger(0)},
	{"adFilterPendingRecords", NewInteger(1)},
	{"adFilterAffectedRecords", NewInteger(2)},
	{"adFilterFetchedRecords", NewInteger(3)},
	{"adFilterConflictingRecords", NewInteger(5)},

	// ── AffectEnum ──────────────────────────────────────────────────────────
	{"adAffectCurrent", NewInteger(1)},
	{"adAffectGroup", NewInteger(2)},
	{"adAffectAll", NewInteger(3)},
	{"adAffectAllChapters", NewInteger(4)},

	// ── RecordStatusEnum ────────────────────────────────────────────────────
	{"adRecordUnmodified", NewInteger(0)},
	{"adRecordModified", NewInteger(1)},
	{"adRecordNew", NewInteger(2)},
	{"adRecordDeleted", NewInteger(4)},

	// ── StreamTypeEnum (ADODB.Stream) ────────────────────────────────────────
	{"adTypeBinary", NewInteger(1)},
	{"adTypeText", NewInteger(2)},

	// ── StreamReadEnum ──────────────────────────────────────────────────────
	{"adReadLine", NewInteger(-2)},
	{"adReadAll", NewInteger(-1)},

	// ── SeekEnum ────────────────────────────────────────────────────────────
	{"adSeekFirstHit", NewInteger(1)},
	{"adSeekAfterHit", NewInteger(2)},
	{"adSeekAfter", NewInteger(4)},
	{"adSeekBeforeHit", NewInteger(8)},
	{"adSeekBefore", NewInteger(16)},
	{"adSeekLastBeforeHit", NewInteger(32)},
	{"adSeekLastHit", NewInteger(64)},

	// ── PositionEnum ────────────────────────────────────────────────────────
	{"adPosBOF", NewInteger(-2)},
	{"adPosEOF", NewInteger(-3)},
	{"adPosUnknown", NewInteger(-1)},

	// ── GetRowsOptionEnum ───────────────────────────────────────────────────
	{"adGetRowsRest", NewInteger(-1)},

	// ── ResyncEnum ──────────────────────────────────────────────────────────
	{"adResyncUnderlyingValues", NewInteger(1)},
	{"adResyncAllValues", NewInteger(2)},

	// ── SearchDirectionEnum ─────────────────────────────────────────────────
	{"adSearchForward", NewInteger(1)},
	{"adSearchBackward", NewInteger(-1)},

	// ── SaveOptionsEnum ─────────────────────────────────────────────────────
	{"adSaveCreateNotExist", NewInteger(1)},
	{"adSaveCreateOverWrite", NewInteger(2)},

	// ── LineSeparatorEnum ───────────────────────────────────────────────────
	{"adCRLF", NewInteger(-1)},
	{"adLF", NewInteger(10)},
	{"adCR", NewInteger(13)},

	// ── ConnectOptionEnum ───────────────────────────────────────────────────
	{"adConnectUnspecified", NewInteger(-1)},
	{"adAsyncConnect", NewInteger(16)},

	// ── SchemaEnum ──────────────────────────────────────────────────────────
	{"adSchemaColumns", NewInteger(4)},
	{"adSchemaForeignKeys", NewInteger(27)},
	{"adSchemaIndexes", NewInteger(12)},
	{"adSchemaProcedures", NewInteger(16)},
	{"adSchemaTables", NewInteger(20)},
	{"adSchemaViews", NewInteger(23)},

	// ── ErrorValueEnum: common ADODB error codes ────────────────────────────
	{"adErrInvalidArgument", NewInteger(3001)},
	{"adErrNoCurrentRecord", NewInteger(3021)},
	{"adErrIllegalOperation", NewInteger(3219)},
	{"adErrInTransaction", NewInteger(3246)},
	{"adErrFeatureNotAvailable", NewInteger(3251)},
	{"adErrItemNotFound", NewInteger(3265)},
	{"adErrObjectInCollection", NewInteger(3367)},
	{"adErrObjectNotSet", NewInteger(3420)},
	{"adErrDataConversion", NewInteger(3421)},
	{"adErrObjectClosed", NewInteger(3704)},
	{"adErrObjectOpen", NewInteger(3705)},
	{"adErrProviderNotFound", NewInteger(3706)},
	{"adErrBoundToCommand", NewInteger(3707)},
	{"adErrInvalidParamInfo", NewInteger(3708)},
	{"adErrInvalidConnection", NewInteger(3709)},
	{"adErrNotReentrant", NewInteger(3710)},
	{"adErrStillExecuting", NewInteger(3711)},
	{"adErrOperationCancelled", NewInteger(3712)},
	{"adErrStillConnecting", NewInteger(3713)},
	{"adErrNotExecuting", NewInteger(3715)},
	{"adErrUnsafeOperation", NewInteger(3716)},
}

// ─── Microsoft Scripting Runtime (FSO) Constants ──────────────────────────────
// Source: Microsoft Scripting Runtime 5.6 Type Library (scrrun.dll).
// Covers FileSystemObject, TextStream, File, Folder, Drive constants.
var fsoTypeLibConstants = []VBSConstant{

	// ── IOMode: used by OpenTextFile / OpenAsTextStream ─────────────────────
	{"ForReading", NewInteger(1)},
	{"ForWriting", NewInteger(2)},
	{"ForAppending", NewInteger(8)},

	// ── Tristate: Unicode/ASCII flag for text streams ───────────────────────
	{"TristateUseDefault", NewInteger(-2)},
	{"TristateTrue", NewInteger(-1)},
	{"TristateFalse", NewInteger(0)},

	// ── FileAttribute ───────────────────────────────────────────────────────
	{"Normal", NewInteger(0)},
	{"ReadOnly", NewInteger(1)},
	{"Hidden", NewInteger(2)},
	{"System", NewInteger(4)},
	{"Volume", NewInteger(8)},
	{"Directory", NewInteger(16)},
	{"Archive", NewInteger(32)},
	{"Alias", NewInteger(64)},
	{"Compressed", NewInteger(128)},

	// ── DriveType ───────────────────────────────────────────────────────────
	{"UnknownType", NewInteger(0)},
	{"Removable", NewInteger(1)},
	{"Fixed", NewInteger(2)},
	{"Remote", NewInteger(3)},
	{"CDRom", NewInteger(4)},
	{"RamDisk", NewInteger(5)},

	// ── StandardStreamTypes ─────────────────────────────────────────────────
	{"StdinStream", NewInteger(0)},
	{"StdoutStream", NewInteger(1)},
	{"StderrStream", NewInteger(2)},

	// ── SpecialFolderConst ──────────────────────────────────────────────────
	{"WindowsFolder", NewInteger(0)},
	{"SystemFolder", NewInteger(1)},
	{"TemporaryFolder", NewInteger(2)},
}

// ─── CDO / CDONTS / CDOSYS Constants ──────────────────────────────────────────
// Covers both CDONTS (NT 4.0) and CDOSYS (Windows 2000+) as used in Classic ASP.
// Source: CDONTS 1.2 / CDO for Windows 2000 (CDOSYS) documentation.
var cdoTypeLibConstants = []VBSConstant{

	// ── CdoMailFormat (CDONTS & CDOSYS) ────────────────────────────────────
	{"CdoMailFormatMime", NewInteger(0)},
	{"CdoMailFormatText", NewInteger(1)},

	// ── CdoBodyFormat (CDONTS) ──────────────────────────────────────────────
	{"CdoBodyFormatHTML", NewInteger(0)},
	{"CdoBodyFormatText", NewInteger(1)},

	// ── CdoPriority / CdoImportance (CDONTS) ───────────────────────────────
	// Note: CDONTS uses 0=Low, 1=Normal, 2=High (different from CDOSYS).
	{"CdoLow", NewInteger(0)},
	{"CdoNormal", NewInteger(1)},
	{"CdoHigh", NewInteger(2)},

	// ── CdoSendUsing (CDOSYS) ───────────────────────────────────────────────
	{"cdoSendUsingPickup", NewInteger(1)},
	{"cdoSendUsingPort", NewInteger(2)},
	{"cdoSendUsingExchange", NewInteger(3)},

	// ── CdoSMTPAuthentication (CDOSYS) ──────────────────────────────────────
	{"cdoAnonymous", NewInteger(0)},
	{"cdoBasic", NewInteger(1)},
	{"cdoNTLM", NewInteger(2)},

	// ── CdoNNTPProcessing (CDOSYS) ──────────────────────────────────────────
	{"cdoNNTPUsePickup", NewInteger(1)},
	{"cdoNNTPUsePort", NewInteger(2)},

	// ── CdoProtocols (CDOSYS) ───────────────────────────────────────────────
	{"cdoProtSMTP", NewInteger(25)},
	{"cdoProtNNTP", NewInteger(119)},
	{"cdoProtPOP3", NewInteger(110)},
	{"cdoProtIMAP", NewInteger(143)},
	{"cdoProtHTTP", NewInteger(80)},

	// ── CdoConnectMode (CDOSYS) ─────────────────────────────────────────────
	{"cdoConnectSameProcess", NewInteger(0)},

	// ── CdoEncoding (CDOSYS) ────────────────────────────────────────────────
	{"cdoEncodingBase64", NewInteger(0)},
	{"cdoEncodingQP", NewInteger(1)},
	{"cdoEncodingUUEncode", NewInteger(2)},
	{"cdoEncodingSevenBit", NewInteger(3)},
	{"cdoEncodingEightBit", NewInteger(4)},
	{"cdoEncodingBinary", NewInteger(5)},
	{"cdoEncodingMacBinhex", NewInteger(6)},

	// ── CdoMHTMLFlags (CDOSYS) ──────────────────────────────────────────────
	{"cdoSuppressAll", NewInteger(0)},
	{"cdoSendTextPlain", NewInteger(1)},
	{"cdoRenderMessage", NewInteger(2)},
	{"cdoSendImages", NewInteger(4)},
	{"cdoIgnoreBodyText", NewInteger(8)},

	// ── CdoPriority (CDOSYS) — uses different values than CDONTS ───────────
	// cdoLow/cdoNormal/cdoHigh already registered above via CDONTS names;
	// the CDOSYS lowercase aliases are below for scripts that use lowercase cdo prefix.
	{"cdoLow", NewInteger(-1)},
	{"cdoNormal", NewInteger(0)},
	{"cdoHigh", NewInteger(1)},

	// ── CDONTS NewMail object constants ─────────────────────────────────────
	{"CdoTo", NewInteger(1)},
	{"CdoCc", NewInteger(2)},
	{"CdoBcc", NewInteger(3)},
}
