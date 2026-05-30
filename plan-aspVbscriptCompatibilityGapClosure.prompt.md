## Plan: ASP VBScript Compatibility Gap Closure

This plan targets full Classic ASP compatibility for VBScript execution paths only (language/runtime + ASP intrinsic object behavior). Discovery confirms core runtime is broad, but key parity gaps remain in Request collection semantics, ServerVariables/ClientCertificate completeness, For Each error semantics, and partial VBScript built-in coverage (notably StrConv modes). Approach: lock failing-compatibility tests first, then patch VM/host dispatch and built-ins to match Classic ASP behavior.

**Steps**
1. Phase 1 (Compatibility Baseline Tests): Add focused regression tests for currently missing semantics before code changes. Include Request.QueryString multi-value indexing, Request default-member missing-key behavior, For Each invalid enumerable errors, full ServerVariables baseline assertions, and ClientCertificate population expectations.
2. Phase 2 (Request Collection Semantics): Normalize Request collection behavior to Classic ASP rules for QueryString/Form/Cookies default and nested access, preserving existing BinaryRead/Form mutual exclusion behavior.
3. Phase 3 (Enumeration/Error Semantics): Update For Each enumerable normalization so invalid targets raise VBScript-compatible runtime errors instead of silent empty iteration.
4. Phase 4 (Intrinsic Object Completeness): Expand Request.ServerVariables and Request.ClientCertificate population from HTTP/TLS context and validate Count/Key/iteration consistency.
5. Phase 5 (VBScript Built-in Parity): Complete StrConv mode coverage for defined VB constants (`vbWide`, `vbNarrow`, `vbKatakana`, `vbHiragana`, `vbFromUnicode`) or explicitly align unsupported modes with documented Classic behavior and errors.
6. Phase 6 (Verification Matrix): Run unit tests and targeted ASP page tests for Request/Response/Session/Application/Server iteration and edge cases across HTTP/CLI/FastCGI where applicable.

**Relevant files**
- `e:/lucas/Desktop/Sites/LGGM-TCP/modules/image/ASP/axonasp2/axonvm/builtins.go` — `vbsAxonEnumValues` For Each normalization and error behavior.
- `e:/lucas/Desktop/Sites/LGGM-TCP/modules/image/ASP/axonasp2/axonvm/compiler_stmt_parser.go` — For Each lowering contract (`__AXON_ENUM_VALUES`, count/item helper path).
- `e:/lucas/Desktop/Sites/LGGM-TCP/modules/image/ASP/axonasp2/axonvm/vm.go` — Request/Response/Server/Session/Application native dispatch semantics; collection member/property behavior.
- `e:/lucas/Desktop/Sites/LGGM-TCP/modules/image/ASP/axonasp2/axonvm/asp/request.go` — RequestCollection key/value/index behavior and lazy parsing order.
- `e:/lucas/Desktop/Sites/LGGM-TCP/modules/image/ASP/axonasp2/server/web_host.go` — population of Request.ServerVariables and Request.ClientCertificate from inbound request/TLS.
- `e:/lucas/Desktop/Sites/LGGM-TCP/modules/image/ASP/axonasp2/axonvm/builtins_vbscript_compat.go` — StrConv compatibility modes.
- `e:/lucas/Desktop/Sites/LGGM-TCP/modules/image/ASP/axonasp2/axonvm/vm_request_test.go` — Request call/collection semantics tests.
- `e:/lucas/Desktop/Sites/LGGM-TCP/modules/image/ASP/axonasp2/axonvm/vm_session_test.go` — Session enumeration/contents parity tests.
- `e:/lucas/Desktop/Sites/LGGM-TCP/modules/image/ASP/axonasp2/axonvm/vm_application_test.go` — Application enumeration/contents parity tests.
- `e:/lucas/Desktop/Sites/LGGM-TCP/modules/image/ASP/axonasp2/axonvm/vm_server_test.go` — Server.Execute/Transfer/GetLastError parity coverage.

**Verification**
1. Add and run targeted Go tests for each discovered gap in `axonvm/*_test.go` with exact expected Classic ASP output/error numbers.
2. Execute focused runtime ASP pages under `www/tests/` for Request collections, For Each invalid targets, and ServerVariables/ClientCertificate exposure.
3. Validate no regressions in existing Request/Form/BinaryRead tests and Session/Application enumeration tests.
4. Verify parity on HTTP path first, then ensure CLI/FastCGI behavior remains consistent for VBScript execution semantics.

**Decisions**
- Included scope: VBScript in Classic ASP runtime semantics, ASP intrinsic object behavior, and enumeration/iteration semantics.
- Excluded scope: JScript engine compatibility and non-ASP desktop interactivity (e.g., MsgBox/InputBox in server mode, which is intentionally unsupported in ASP).
- Compatibility priority order: (1) correctness of observable behavior and error semantics, (2) parity of collection/object APIs, (3) preservation of existing performance-oriented architecture.

**Further Considerations**
1. Decide whether unsupported StrConv modes should emulate best-effort transforms or raise explicit VBScript-compatible runtime errors when exact conversion is unavailable.
2. Decide whether Request key casing should preserve original incoming case or remain normalized/lowercase during enumeration when matching IIS output expectations.
3. Decide whether Response.Cookies should expose additional collection-like metadata (`Count`/`Key`) or remain minimal if Classic ASP does not guarantee those members on Response-side cookies.
