# Architectural Plan: Compatibility & Engine Patches

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
  3. You **must ensure** that the formatting functions respect locale settings (e.g., currency symbol, decimal separator) and handle edge cases like negative zero, scientific notation, and rounding according to **Microsoft VBScript** rules.

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
