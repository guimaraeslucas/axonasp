# Architectural Plan: Compatibility & Engine Patches

## 1. VBScript For Loop Numeric Coercion
- **Affected Files:**
  - `axonvm/compiler_stmt_parser.go` (`parseForToStatement`)
  - `axonvm/vm.go` (`OpForNextFastInt`, `OpForNextFastGlobalInt`, comparison operators)
  - `axonvm/opcode.go` (`OpCoerceToNumeric` or boundary initialization)
- **Defect:**
  In VBScript, `For` loop boundaries provided as strings (e.g. `limit = "5"` or `For i = 1 To "5"`) must be coerced to numeric subtypes (`VTInteger` / `VTDouble`) before loop evaluation. Currently, non-coerced boundaries cause string lexicographical comparisons in `OpLte`/`OpGte` or bypass the fast integer loop paths.
- **Implementation Strategy:**
  1. Introduce opcode `OpCoerceToNumeric` or enforce numeric conversion right after evaluating the `start`, `limit`, and `step` expressions in `parseForToStatement`.
  2. Emit numeric coercion after evaluating `initExpr`, `limitExpr`, and `stepExpr` before assigning them to their respective storage slots (`loopVarName`, `endName`, `stepName`).
  3. In `axonvm/vm.go`, implement `OpCoerceToNumeric` (or within boundary setup):
     - If TOS is `VTInteger` or `VTDouble`, keep as-is.
     - If TOS is `VTString`, parse via `parseVBScriptNumber` / `vm.asFloat` / `strconv.ParseInt`. If integer, push `VTInteger`; if floating point, push `VTDouble`. If unparseable, raise `vbscript.TypeMismatch`.
     - If TOS is `VTNull`, raise `vbscript.TypeMismatch` or propagate according to VBScript semantics.
  4. Ensure `OpForNextFastInt` and `OpForNextFastGlobalInt` fast paths are taken when coerced values are integer, preventing unnecessary float or string fallbacks.

## 2. FormatCurrency Type Handling
- **Affected Files:**
  - `axonvm/builtins_vbscript_compat.go` (`vbsCompatFormatCurrency`, `vbsCompatFormatPercent`)
- **Defect:**
  `vbsCompatFormatCurrency` extracts `value := float64(args[0].Num)`. When `args[0]` is a `VTString`, `args[0].Num` is `0`, causing numeric strings (e.g. `"1234.50"`) to format as zero (`$0.00`).
- **Implementation Strategy:**
  1. In `vbsCompatFormatCurrency`, inspect `args[0].Type`:
     - `VTInteger`: `value = float64(args[0].Num)`
     - `VTDouble`: `value = args[0].Flt`
     - `VTString`: Convert string to float using `strconv.ParseFloat` after trimming whitespace. If invalid numeric format, raise `vbscript.TypeMismatch`.
     - `VTNull`: Return `NewNull()` (matching Microsoft VBScript `FormatCurrency(Null)` behavior).
     - `VTEmpty`: Value evaluates to `0.0`.
  2. Apply the identical type coercion pipeline to `vbsCompatFormatPercent` and `vbsCompatFormatNumber`.

## 3 & 4. Forward References & Hoisting in Single-Pass Compiler - Done
- **Affected Files:**
  - `axonvm/compiler.go` (`Compiler`, `patchForwardConstSites`, `registerForwardConstPatch`)
  - `axonvm/compiler_expr.go` (`parseIdentifierOrCallExpr`)
  - `axonvm/compiler_stmt_parser.go` (`parseProcedureParameterNames`, `parseSubFunction`, `parseConstStatement`)
  - `axonvm/opcode.go` (`OpResolveDynamic`)
  - `axonvm/vm.go` (`OpResolveDynamic` execution, `opcodeOperandSize`, `remapExecuteGlobalBytecode`)
- **Defect:**
  1. Constants declared after procedure definitions resolve to uninitialized local frame slots (`OpGetLocal`) inside procedures because `patchForwardConstSites` only tracks `OpGetGlobal`.
  2. Bare parameterless function calls placed before the `Function` declaration (e.g. `x = myFunc + 1`) are emitted as variable reads (`OpGetLocal` or `OpGetGlobal`) and evaluate to `Empty`.
- **Implementation Strategy:**
  1. Single-Pass Constraint: Strictly single-pass (no AST allowed for VBScript). Late-binding resolution is required for unknown identifiers.
  2. Opcode Addition:
     - Register `OpResolveDynamic` (primary opcode or extended opcode `ExtOpResolveDynamic` depending on opcode availability).
     - Operand: 2 bytes `[uint16 constNameIdx]` referencing the identifier string constant.
     - Synchronize the Three-Way Contract:
       - Execution loop in `axonvm/vm.go`
       - `opcodeOperandSize` in `axonvm/vm.go`
       - `remapExecuteGlobalBytecode` in `axonvm/vm.go`
       - `OpCode.String()` in `axonvm/opcode.go`
  3. Compiler Emission:
     - When encountering an unresolved identifier in expression contexts (where it is neither a known local, declared global, constant, nor registered procedure):
       - If `Option Explicit` is off, or if within a procedure where it could refer to a forward-declared global constant or procedure:
         - Emit `OpResolveDynamic [nameIdx]`.
  4. VM Execution (`OpResolveDynamic`):
     - Lookup hierarchy at runtime:
       1. Global constants (`vm.globalConstants` / `vm.constants`).
       2. Global variable table (`vm.Globals`).
       3. Function/Procedure registry (`vm.functions` / `UserSub`).
     - If resolved target is a `VTUserSub`:
       - If `UserSubIsFunc()` and `UserSubParamCount() == 0`:
         - Automatically invoke function call via `vm.beginUserSubCall(target, nil, false, 0)`.
       - Else, push the `VTUserSub` value.
     - If resolved to variable or constant, push its value.
     - If not found: push `NewEmpty()` (or raise `vbscript.VariableUndefined` if `Option Explicit` is active).

## 5. ADODB Identity Retrieval
- **Affected Files:**
  - `axonvm/lib_adodb.go` (`adodbExecWriteback`, `adodbPersistRecordsetInsert`, `adodbRecordsetAddNew`)
- **Defect:**
  `Recordset.Update` after `AddNew` fails to populate auto-generated identity values on database drivers (notably SQL Server / `go-mssqldb` where `Result.LastInsertId()` is unsupported and returns an error).
- **Implementation Strategy:**
  1. In `adodbExecWriteback`:
     - If `conn.db != nil` and `wantIdentity == true`:
       - Check driver / provider type. If driver does not support `LastInsertId()` (or if `res.LastInsertId()` errors):
         - Execute `SELECT SCOPE_IDENTITY() AS NewID` (or provider-specific identity query: `SELECT @@IDENTITY` for SQL Server, `SELECT LAST_INSERT_ID()` for MySQL, `SELECT last_insert_rowid()` for SQLite) on the identical connection/transaction context (`conn.tx` or `conn.dbConn`).
         - Scan scalar identity integer into `insertID`.
  2. In `adodbPersistRecordsetInsert`:
     - Detect identity column from recordset fields (e.g. `ID`, `iid`, or columns marked as identity/auto-increment).
     - If identity column is present in the table schema or recordset columns, capture `insertID`.
     - Write the generated identity into `row[strings.ToLower(keyField)] = NewInteger(insertID)`.
     - Update `rs.data[rs.currentRow]` and refresh field proxy cache so subsequent reads (`rs.Fields("id").Value`) immediately reflect the new identity.

## 6. ADODB Single Column Updatability
- **Affected Files:**
  - `axonvm/lib_adodb.go` (`adodbResolveUpdatableSource`, `adodbSimpleSelectSourceRE`, `adodbPersistRecordsetUpdate`)
- **Defect:**
  Updating a recordset opened with a single-column projection (e.g. `SELECT Name FROM Customers WHERE ID = 1`) fails with:
  `"ADODB.Recordset.Update: source is not one simple updatable single-table query"` because `adodbSimpleSelectSourceRE` strictly requires `SELECT * FROM`.
- **Implementation Strategy:**
  1. Replace `adodbSimpleSelectSourceRE` with a generalized regex:
     - Pattern: `(?is)^\s*select\s+(.+?)\s+from\s+(.+?)(?:\s+where\s+(.+?))?(?:\s+order\s+by\s+.+)?\s*$`
  2. Parse the table expression: ensure single base table without joins or subqueries.
  3. Extract target table name and existing `WHERE` clause.
  4. In `adodbPersistRecordsetUpdate`:
     - Loosen key column requirement: if a `WHERE` clause was captured from the source SQL, utilize that `WHERE` clause for the `UPDATE` statement target.
     - Build `SET` statement solely from mutated columns (`rs.pendingUpdateFields`), permitting updates even when the primary key column was excluded from the `SELECT` column list.

## 7. Implicit ByRef Variable Allocation
- **Affected Files:**
  - `axonvm/compiler_expr.go` (`rewriteLastVarLoadToArgRef`)
  - `axonvm/vm.go` (`collectByRefsAndUnwrap`, `applyByRefWritebacksFromArgs`, `OpRet`)
- **Defect:**
  Implicitly declared variables passed ByRef to procedures returning objects/recordsets lose their reference because the compiler disables ByRef rewrite for undeclared variables (`!c.declaredGlobals[name]`), passing them ByVal. Additionally, `VTNativeObject` assignments are not written back properly to caller scopes.
- **Implementation Strategy:**
  1. In `compiler_expr.go` (`rewriteLastVarLoadToArgRef`):
     - Remove the restrictive check `if !c.declaredGlobals[name] { return }`.
     - If the variable is undeclared, allocate it dynamically in caller scope (registering as implicit global or local) and emit `OpArgGlobalRef` or `OpArgLocalRef`.
  2. In `axonvm/vm.go` (`OpRet` & `applyByRefWritebacksFromArgs`):
     - When callee updates a ByRef argument with a `VTNativeObject` (or `VTObject`), write the object value directly into the caller's target slot (`vm.Globals[wb.callerIdx]` or `vm.stack[callerSlot]`).
     - Correctly manage reference tracking: increment ref count on new assignment and decrement on overwritten slot without dropping native object pointer mappings.

## 8. Dictionary Cookie Persistence
- **Affected Files:**
  - `axonvm/asp/response.go` (`SetCookieSubKey`, `ResponseCookie`)
  - `axonvm/asp/request.go` (`parseCookieSubKeys`)
  - `axonvm/vm.go` (`dispatchNativeCall` / `responseCookieItems`)
- **Defect:**
  Setting a dictionary cookie subkey (e.g. `Response.Cookies("MGSsession")("key2") = val`) overwrites existing subkeys previously received from the client in `Request.Cookies("MGSsession")`.
- **Implementation Strategy:**
  1. In `axonvm/asp/response.go` (`SetCookieSubKey`):
     - When a cookie entry does not yet exist in `r.cookies` (or has an empty `Value`):
       - Inspect the incoming HTTP request (`r.req`). If the cookie exists in incoming HTTP headers (`r.req.Cookie(cookieName)`):
         - Load incoming raw cookie value.
         - Parse existing subkeys using `parseCookieSubKeys`.
  2. Merge the newly assigned subkey into the parsed dictionary: `subKeys[strings.ToLower(subKey)] = value`.
  3. Rebuild and serialize into `cookie.Value = "key1=val1&key2=val2"` format.
  4. Ensure subkey persistence across repeated assignments in the same request context.

---

## Verification & Validation Protocol
1. **Compilation Check:** Compile `./axonasp-http.exe` and `./axonasp-cli.exe` via `go build`.
2. **Go Unit Tests:** Run `go test ./axonvm/...` to verify stack integrity, opcode synchronization, and zero regressions.
3. **Regression Guard:** Verify that JScript AST parsing remains strictly isolated in `./jscript/` and zero-allocation stack discipline in VBScript VM is preserved.
