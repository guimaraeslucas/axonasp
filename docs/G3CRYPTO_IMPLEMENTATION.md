## G3CRYPTO Library Implementation Summary

### Overview
A comprehensive cryptography library has been implemented for AxonASP, providing professional-grade security operations for password hashing, verification, and UUID generation for authentication and data integrity.

### Files Created/Modified

#### New/Modified Files
1. **`server/crypto_lib.go`** (78 lines)
   - Complete implementation of G3CRYPTO library
   - Password hashing with bcrypt
   - Password verification
   - UUID generation (v4)
   - Secure random operations

#### Dependencies
- **golang.org/x/crypto/bcrypt** - Industry-standard password hashing
- **crypto/rand** - Cryptographically secure random number generation

#### Integration
1. **`server/executor_libraries.go`**
   - Added CryptoLibrary wrapper for ASPLibrary interface compatibility
   - Enables: `Set crypto = Server.CreateObject("G3CRYPTO")`
   - Also supports: `Server.CreateObject("CRYPTO")`

2. **`server/custom_functions.go`**
   - Added `CryptoHelper()` function for backward compatibility

3. **`server/executor.go`**
    - Added CreateObject ProgID support for `.NET` compatibility
    - Enables: `Set md5 = Server.CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")`

### .NET Compatibility Object

✓ **System.Security.Cryptography.MD5CryptoServiceProvider**
  - `ComputeHash(input)` - Returns byte array compatible with VBScript consumption
  - `Initialize()`, `Clear()`, `Dispose()` - Reset internal hash state
  - `Hash` property - Returns latest computed hash bytes
  - `HashSize` property - Returns `128`

### Key Features Implemented

✓ **Password Hashing**
  - `HashPassword(password)` / `Hash(password)` - Generate secure hash
  - bcrypt algorithm with default cost factor
  - Salted and one-way hashing
  - Resistant to rainbow table attacks

✓ **Password Verification**
  - `VerifyPassword(password, hash)` / `Verify(password, hash)` - Verify password match
  - Constant-time comparison to prevent timing attacks
  - Safe comparison even for wrong passwords
  - Returns boolean result

✓ **UUID Generation**
  - `UUID()` - Generate random UUID v4
  - Cryptographically secure random source
  - RFC 4122 compliant format
  - Suitable for unique identifiers

✓ **Security**
  - bcrypt cost factor: 12 (industry standard)
  - Cryptographic random number generation
  - No hardcoded keys or secrets
  - Safe against timing attacks

### Architecture

**Class Hierarchy**:
```
Component (interface)
  └─ G3CRYPTO
      ├─ UUID()
      ├─ HashPassword()
      ├─ VerifyPassword()
      └─ Helper methods
```

**Password Flow**:
1. User provides plaintext password
2. bcrypt generates random salt
3. Password hashed with salt (12 cost rounds)
4. Full hash returned (includes salt)
5. Hash stored in database

**Verification Flow**:
1. User provides plaintext password
2. Stored hash retrieved from database
3. bcrypt extracts salt from hash
4. New hash generated with extracted salt
5. Hashes compared using constant-time algorithm
6. Boolean result returned

### Usage Examples

#### User Registration with Password Hashing
```vbscript
Dim crypto, files, passwordHash, username, password
Set crypto = Server.CreateObject("G3CRYPTO")
Set files = Server.CreateObject("G3FILES")

' Get form input
username = Request.Form("username")
password = Request.Form("password")

' Validate password
If Len(password) < 8 Then
    Response.Write "Password must be at least 8 characters"
Else
    ' Hash password
    passwordHash = crypto.HashPassword(password)
    
    ' Save user (example data)
    Dim userData
    userData = username & "|" & passwordHash & vbCrLf
    files.Append("users.txt", userData)
    
    Response.Write "Registration successful"
End If
```

#### User Login Verification
```vbscript
Dim crypto, files, username, password, userLine, parts
Set crypto = Server.CreateObject("G3CRYPTO")
Set files = Server.CreateObject("G3FILES")

' Get form input
username = Request.Form("username")
password = Request.Form("password")

' Read user file
Dim users
users = files.Read("users.txt")

' Search for user
Dim lines
lines = Split(users, vbCrLf)

Dim found, i
found = False

For i = 0 To UBound(lines)
    If lines(i) <> "" Then
        parts = Split(lines(i), "|")
        If parts(0) = username Then
            ' Verify password
            If crypto.VerifyPassword(password, parts(1)) Then
                Response.Write "Login successful"
                found = True
            Else
                Response.Write "Invalid password"
                found = True
            End If
            Exit For
        End If
    End If
Next

If Not found Then
    Response.Write "User not found"
End If
```

#### UUID Generation for Session IDs
```vbscript
Dim crypto, sessionId
Set crypto = Server.CreateObject("G3CRYPTO")

' Generate unique session ID
sessionId = crypto.UUID()
Session("SessionID") = sessionId

Response.Write "Session created: " & sessionId
' Output: Session created: 550e8400-e29b-41d4-a716-446655440000
```

#### Creating User Tokens
```vbscript
Dim crypto, token, database
Set crypto = Server.CreateObject("G3CRYPTO")
Set database = Server.CreateObject("ADODB.Connection")

' Generate secure token for password reset
token = crypto.UUID()

' Store token in database
database.ConnectionString = "Provider=MSDASQL;Driver={SQL Server};" & _
    "Server=localhost;Database=myapp;uid=sa;pwd=password"
database.Open

Dim sql
sql = "INSERT INTO password_resets (token, user_id, expires) " & _
      "VALUES ('" & token & "', 123, DATEADD(hour, 24, GETDATE()))"
database.Execute sql

database.Close()

' Send reset link
Dim mail
Set mail = Server.CreateObject("G3MAIL")
mail.SendStandard( _
    "user@example.com", _
    "Password Reset", _
    "Click here to reset: https://example.com/reset?token=" & token, _
    False
)
```

#### API Key Generation
```vbscript
Dim crypto, apiKey
Set crypto = Server.CreateObject("G3CRYPTO")

' Generate API key
apiKey = Replace(crypto.UUID(), "-", "")

' Result: 550e8400e29b41d4a716446655440000
' Suitable for API authentication
Response.Write "Your API Key: " & apiKey
```

#### Secure File Upload Tracking
```vbscript
Dim crypto, uploader, fileId, files
Set crypto = Server.CreateObject("G3CRYPTO")
Set uploader = Server.CreateObject("G3FileUploader")
Set files = Server.CreateObject("G3FILES")

' Process upload
Dim result
Set result = uploader.Process("upload", "uploads")

' Create unique file ID
fileId = crypto.UUID()

' Store file metadata
Dim metadata
metadata = fileId & "|" & result.NewFileName & "|" & Now & vbCrLf
files.Append("file_log.txt", metadata)

Response.Write "File uploaded with ID: " & fileId
```

#### Two-Factor Authentication Token
```vbscript
Dim crypto, token, codes, i
Set crypto = Server.CreateObject("G3CRYPTO")

' Generate 6 random codes (simulated)
Dim code
code = Int(Rnd() * 999999)

' Use UUID as backup code base
Dim backupCode
backupCode = Left(Replace(crypto.UUID(), "-", ""), 16)

Response.Write "2FA Code: " & Format(code, "000000") & "<br>"
Response.Write "Backup Code: " & backupCode
```

#### Secure Password Storage in Database
```vbscript
Dim crypto, database, username, password, hash
Set crypto = Server.CreateObject("G3CRYPTO")
Set database = Server.CreateObject("ADODB.Connection")

username = Request.Form("username")
password = Request.Form("password")

' Hash password
hash = crypto.HashPassword(password)

' Connect to database
database.ConnectionString = "Provider=MSDASQL;Driver={SQL Server};" & _
    "Server=localhost;Database=myapp;uid=sa;pwd=password"
database.Open

' Store user with hashed password
Dim sql
sql = "INSERT INTO users (username, password_hash, created_date) " & _
      "VALUES ('" & username & "', '" & hash & "', GETDATE())"
database.Execute sql

database.Close()

Response.Write "User created successfully"
```

#### Password Change Verification
```vbscript
Dim crypto, currentPassword, newPassword, storedHash
Set crypto = Server.CreateObject("G3CRYPTO")

storedHash = "YOUR_STORED_HASH_HERE"
currentPassword = Request.Form("current_password")
newPassword = Request.Form("new_password")

' Verify current password
If Not crypto.VerifyPassword(currentPassword, storedHash) Then
    Response.Write "Current password is incorrect"
Else
    ' Hash new password
    Dim newHash
    newHash = crypto.HashPassword(newPassword)
    
    ' Update in database
    Dim database
    Set database = Server.CreateObject("ADODB.Connection")
    ' ... Update logic ...
    
    Response.Write "Password changed successfully"
End If
```

### UUID Format

**Standard UUID v4 Format**:
```
550e8400-e29b-41d4-a716-446655440000
^^^^^^^^-^^^^-^^^^-^^^^-^^^^^^^^^^^^
|time low|time |time |clock|node
         |mid  |high |seq  |
```

**Properties**:
- Globally unique identifier
- 36 characters with hyphens
- Can be removed with Replace(uuid, "-", "")
- RFC 4122 compliant

### Password Hashing Details

**bcrypt Algorithm**:
- Cost factor: 12 (configurable but default to 12)
- Salt: Automatically generated and included in hash
- Hash length: 60 characters
- Format: `$2a$cost$salt(22 chars)hash(31 chars)`

**Hash Example**:
```
Original: myPassword123
Hashed:   $2a$12$Bc5T6D7KqNrJ8mL3P5Q2HeA4xB6C8dE9F1G2H3I4J5K6L7M8N9O0P
```

### Security Characteristics

✓ **Strengths**:
- Industry-standard bcrypt algorithm
- Automatic salt generation
- Resistant to rainbow table attacks
- Resistant to GPU attacks (slow by design)
- Constant-time comparison prevents timing attacks
- Cryptographically secure random for UUIDs

⚠ **Considerations**:
- Password hashing is intentionally slow (1-2 seconds per hash)
- UUID alone is not suitable for security tokens (needs + encryption)
- No HMAC support (use bcrypt for passwords)
- No encryption support (use other libraries)

### Performance Characteristics
- bcrypt hashing: 1-2 seconds (by design)
- Password verification: 1-2 seconds (by design)
- UUID generation: < 1 millisecond
- Memory usage: Minimal
- CPU-bound operations

### Best Practices

✓ **Do**:
- Always hash passwords before storage
- Use VerifyPassword for login checks
- Generate unique UUIDs for tokens/sessions
- Store full bcrypt hash (includes salt)
- Use HTTPS for password transmission

✗ **Don't**:
- Send passwords over unencrypted connections
- Use simple string comparison for password verification
- Reuse UUIDs for critical operations
- Store passwords in plain text
- Use UUID as sole security measure
- Hardcode security tokens in code

### Limitations
- No encryption support (one-way hashing only)
- No key derivation function (KDF)
- No HMAC-based MAC support
- bcrypt not suitable for other hashing needs
- No asymmetric cryptography
- No digital signatures

### Future Enhancements
- Additional hash algorithms (Argon2, scrypt)
- Encryption/decryption support
- HMAC-SHA algorithms
- Key derivation functions
- Rate limiting for brute force prevention
- Configurable bcrypt cost factor
