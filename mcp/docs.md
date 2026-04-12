## ServerObject: G3Crypto Function: UUID()
**Keywords:** uuid, guid, unique, identity, random, identifier, rfc4122, v4
**Description:** Generates a version 4 Universally Unique Identifier (UUID) as defined in RFC 4122.
Return: `String` - A 36-character string in the format "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx".
**Observations:** Uses cryptographically secure pseudorandom numbers. The output is always lowercase hexadecimal. This is the preferred method for generating unique tokens or primary keys in AxonASP.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
Dim myID
myID = ax.UUID()
Response.Write "Generated ID: " & myID
```

## ServerObject: G3Crypto Function: HashPassword(password)
**Keywords:** password, hash, bcrypt, security, authentication, salt, blowfish
**Description:** Hashes a plain-text password using the bcrypt adaptive hashing algorithm.
Input: `password` (String) - The plain-text password to hash.
Return: `String` - The encoded bcrypt hash, including the algorithm identifier, cost, and salt.
**Observations:** Bcrypt is inherently resistant to brute-force and rainbow table attacks due to its built-in salt and configurable cost. The cost can be adjusted using `SetBcryptCost`. Returns an empty string if hashing fails.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
Dim hashedPassword
hashedPassword = ax.HashPassword("mySecretPassword")
' Store hashedPassword in your database
```

## ServerObject: G3Crypto Function: VerifyPassword(password, hash)
**Keywords:** password, verify, check, authentication, bcrypt, login
**Description:** Validates a plain-text password against an existing bcrypt hash.
Input: `password` (String) - The plain-text password provided by the user.
Input: `hash` (String) - The bcrypt hash stored previously.
Return: `Boolean` - True if the password matches the hash, False otherwise.
**Observations:** This method is timing-attack resistant. It internally extracts the salt and cost from the hash string to perform the comparison.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
Dim isValid
isValid = ax.VerifyPassword("userInput", storedHashFromDB)
If isValid Then
    Response.Write "Access Granted"
Else
    Response.Write "Access Denied"
End If
```

## ServerObject: G3Crypto Function: SetBcryptCost(cost)
**Keywords:** bcrypt, cost, complexity, security, performance, tuning
**Description:** Configures the computational cost (work factor) for bcrypt hashing operations.
Input: `cost` (Integer) - An integer between 4 and 31. Default is 10.
Return: `Boolean` - True if the cost was successfully updated, False if the value is out of range.
**Observations:** Increasing the cost exponentially increases the time required to compute the hash. A cost of 12 or 14 is recommended for modern servers. Setting it too high can lead to Denial of Service (DoS) due to CPU exhaustion.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
If ax.SetBcryptCost(12) Then
    Response.Write "Bcrypt cost updated to 12"
End If
```

## ServerObject: G3Crypto Function: GetBcryptCost()
**Keywords:** bcrypt, cost, settings, configuration, query
**Description:** Retrieves the current bcrypt cost setting used by the object instance.
Return: `Integer` - The current work factor.
**Observations:** Useful for auditing security settings at runtime.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
Response.Write "Current Bcrypt Cost: " & ax.GetBcryptCost()
```

## ServerObject: G3Crypto Function: RandomBytes(size)
**Keywords:** random, bytes, binary, entropy, security, token
**Description:** Generates a sequence of cryptographically secure random bytes.
Input: `size` (Integer) - The number of bytes to generate.
Return: `Array` (VTArray) - A zero-based array of integers (0-255).
**Observations:** Uses the operating system's secure entropy source (e.g., /dev/urandom or CryptGenRandom). Essential for generating nonces or session keys.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
Dim key, i
key = ax.RandomBytes(16)
For i = 0 To UBound(key)
    Response.Write Hex(key(i))
Next
```

## ServerObject: G3Crypto Function: RandomHex(size)
**Keywords:** random, hex, string, token, entropy, security
**Description:** Generates a sequence of secure random bytes and returns them as a hexadecimal string.
Input: `size` (Integer) - The number of bytes to generate (the output string will be double this length).
Return: `String` - A lowercase hexadecimal string.
**Observations:** A `size` of 16 produces a 32-character hex string.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
Dim token
token = ax.RandomHex(32)
Response.Write "Random Token: " & token
```

## ServerObject: G3Crypto Function: MD5(data)
**Keywords:** md5, hash, digest, integrity, legacy
**Description:** Computes the MD5 hash of the input data and returns a hexadecimal string.
Input: `data` (String or Array) - The data to hash.
Return: `String` - A 32-character hexadecimal string.
**Observations:** MD5 is cryptographically broken and should NOT be used for password storage. Use it only for legacy compatibility or non-security-critical integrity checks.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
Response.Write "MD5: " & ax.MD5("Hello World")
```

## ServerObject: G3Crypto Function: SHA256(data)
**Keywords:** sha256, hash, sha2, digest, security, integrity
**Description:** Computes the SHA-256 hash of the input data and returns a hexadecimal string.
Input: `data` (String or Array) - The data to hash.
Return: `String` - A 64-character hexadecimal string.
**Observations:** SHA-256 is the industry standard for secure hashing and digital signatures.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
Response.Write "SHA256: " & ax.SHA256("Hello World")
```

## ServerObject: G3Crypto Function: SHA256Bytes(data)
**Keywords:** sha256, hash, bytes, binary, digest, security
**Description:** Computes the SHA-256 hash and returns the raw bytes.
Input: `data` (String or Array) - The data to hash.
Return: `Array` (VTArray) - A 32-byte array of integers.
**Observations:** Useful when the hash needs to be used in further cryptographic operations rather than being displayed.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
Dim digest
digest = ax.SHA256Bytes("Secret Message")
```

## ServerObject: G3Crypto Function: HMACSHA256(data, key)
**Keywords:** hmac, sha256, message, authentication, sign, signature, integrity, secret
**Description:** Calculates a Keyed-Hash Message Authentication Code (HMAC) using the SHA-256 hash function.
Input: `data` (String) - The message to sign.
Input: `key` (String) - The secret key used for the HMAC.
Return: `String` - A 64-character hexadecimal string representing the digest.
**Observations:** HMAC provides both integrity and authenticity by requiring a secret key. This is essential for signing webhooks or generating secure tokens (like JWT signatures).
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
Dim signature
signature = ax.HMACSHA256("message_payload", "secret_key")
Response.Write "Signature: " & signature
```

## ServerObject: G3Crypto Function: PBKDF2SHA256(password, salt, iterations, keyLength)
**Keywords:** pbkdf2, kdf, key, derivation, password, security, sha256, stretching
**Description:** Derives a cryptographic key from a password using the PBKDF2-HMAC-SHA256 algorithm.
Input: `password` (String) - The input password or master key.
Input: `salt` (String) - A random salt value to prevent rainbow table attacks.
Input: `iterations` (Integer) - The number of hash iterations (default 100,000).
Input: `keyLength` (Integer) - The desired length of the derived key in bytes (default 32).
Return: `String` - The derived key as a hexadecimal string.
**Observations:** PBKDF2 is a key-stretching algorithm that makes brute-force attacks significantly more difficult.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
Dim derivedKey
derivedKey = ax.PBKDF2SHA256("password123", "random_salt", 100000, 32)
Response.Write "Derived Key: " & derivedKey
```

## ServerObject: G3Crypto Function: ComputeHash(data, algorithm)
**Keywords:** hash, compute, generic, algorithm, md5, sha1, sha256, dynamic
**Description:** Computes a hash using a dynamically specified algorithm.
Input: `data` (String or Array) - The data to hash.
Input: `algorithm` (String) - The algorithm name (e.g., "md5", "sha1", "sha256", "sha512").
Return: `Array` (VTArray) - The raw hash bytes.
**Observations:** This is the most flexible hashing method. If the algorithm is omitted, it uses the default algorithm configured during object instantiation.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Crypto")
Dim h
h = ax.ComputeHash("Some Data", "sha512")
```

## ServerObject: G3Files Function: Exists(path)
**Keywords:** file, exists, directory, check, filesystem, path, verification
**Description:** Determines if a file or directory exists at the specified path.
Input: `path` (String) - The relative or absolute path to check.
Return: `Boolean` - True if the file/directory exists, False otherwise.
**Observations:** The path is resolved relative to the web root or the configured sandbox. This function is a high-performance alternative to `FileSystemObject.FileExists`.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Files")
If ax.Exists("/data/config.json") Then
    Response.Write "File found"
Else
    Response.Write "File not found"
End If
```

## ServerObject: G3Files Function: Read(path, encoding)
**Keywords:** read, file, content, text, encoding, utf8, utf16, ascii
**Description:** Reads the entire contents of a file as a string.
Input: `path` (String) - The path to the file.
Input: `encoding` (String) - Optional. The text encoding (e.g., "utf-8", "utf-16le", "ascii", "iso-8859-1").
Return: `String` - The file content.
**Observations:** Automatically detects Byte Order Marks (BOM) for UTF-8 and UTF-16. If no encoding is provided, it defaults to UTF-8. Returns an empty string if the file cannot be read.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Files")
Dim content
content = ax.Read("notes.txt", "utf-8")
Response.Write "Content: " & content
```

## ServerObject: G3Files Function: Write(path, content, encoding, lineEnding, includeBOM)
**Keywords:** write, save, create, file, content, encoding, eol, bom
**Description:** Creates or overwrites a file with the provided string content.
Input: `path` (String) - The target file path.
Input: `content` (String) - The text to write.
Input: `encoding` (String) - Optional. Target encoding (default "utf-8").
Input: `lineEnding` (String) - Optional. Line ending style ("windows"/"crlf", "linux"/"lf", "mac"/"cr").
Input: `includeBOM` (Boolean) - Optional. Whether to include a Byte Order Mark.
Return: `Boolean` - True if successful, False otherwise.
**Observations:** Automatically creates parent directories if they don't exist. This is a powerful, all-in-one file writing utility.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Files")
Dim success
success = ax.Write("logs/app.log", "App started", "utf-8", "crlf", False)
If success Then
    Response.Write "Log saved"
End If
```

## ServerObject: G3Files Function: Append(path, content, encoding, lineEnding)
**Keywords:** append, add, write, file, log, content
**Description:** Appends string content to the end of an existing file or creates it if it doesn't exist.
Input: `path` (String) - The target file path.
Input: `content` (String) - The text to append.
Input: `encoding` (String) - Optional. Encoding to use (default "utf-8").
Input: `lineEnding` (String) - Optional. Line ending style.
Return: `Boolean` - True if successful.
**Observations:** Unlike `Write`, this does not support `includeBOM` because appending a BOM to the middle of a file would corrupt the encoding.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Files")
ax.Append("counter.txt", "New Line" & vbCrLf)
```

## ServerObject: G3Files Function: Delete(path)
**Keywords:** delete, remove, unlink, file, filesystem
**Description:** Deletes a file from the filesystem.
Input: `path` (String) - The path to the file to delete.
Return: `Boolean` - True on success.
**Observations:** Does not move to trash; deletion is permanent.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Files")
If ax.Delete("temp_file.tmp") Then
    Response.Write "Deleted"
End If
```

## ServerObject: G3Files Function: Copy(source, destination)
**Keywords:** copy, duplicate, file, filesystem
**Description:** Copies a file from the source path to the destination path.
Input: `source` (String) - Source file path.
Input: `destination` (String) - Destination file path.
Return: `Boolean` - True on success.
**Observations:** Overwrites the destination if it already exists.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Files")
ax.Copy "original.txt", "backup.txt"
```

## ServerObject: G3Files Function: Move(source, destination)
**Keywords:** move, rename, relocate, file, filesystem
**Description:** Moves or renames a file.
Input: `source` (String) - Original path.
Input: `destination` (String) - New path.
Return: `Boolean` - True on success.
**Observations:** Can be used to rename files within the same directory or move them across the filesystem.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Files")
ax.Move "old_name.txt", "new_name.txt"
```

## ServerObject: G3Files Function: Size(path)
**Keywords:** size, file, length, bytes, metadata
**Description:** Retrieves the size of a file in bytes.
Input: `path` (String) - The path to the file.
Return: `Integer` (int64) - The file size in bytes.
**Observations:** Returns 0 if the file does not exist or is a directory.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Files")
Response.Write "Size: " & ax.Size("data.zip") & " bytes"
```

## ServerObject: G3Files Function: MkDir(path)
**Keywords:** mkdir, makedir, directory, folder, create
**Description:** Creates a new directory, including any necessary parent directories.
Input: `path` (String) - The directory path to create.
Return: `Boolean` - True on success.
**Observations:** Equivalent to `mkdir -p` in Unix or `Directory.CreateDirectory` in .NET.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Files")
ax.MkDir "/assets/uploads/images/2026"
```

## ServerObject: G3Files Function: List(path)
**Keywords:** list, directory, files, enumeration, folder, scan
**Description:** Returns a list of filenames within a directory.
Input: `path` (String) - The directory path to scan.
Return: `Array` (VTArray) - A zero-based array of strings containing filenames.
**Observations:** Only returns filenames, not full paths. Directories are excluded from the results. Returns an empty array if the directory is empty or inaccessible.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Files")
Dim files, i
files = ax.List("/uploads")
For i = 0 To UBound(files)
    Response.Write "File: " & files(i) & "<br>"
Next
```

## ServerObject: G3Files Function: NormalizeEOL(text, style)
**Keywords:** eol, line, ending, normalize, crlf, lf, string, formatting
**Description:** Converts all line endings in a string to a consistent style.
Input: `text` (String) - The input text.
Input: `style` (String) - The target style ("windows", "linux", "mac").
Return: `String` - The normalized text.
**Observations:** Handles mixed line endings correctly. "windows" uses CRLF, "linux" uses LF, and "mac" uses CR.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Files")
Dim cleanText
cleanText = ax.NormalizeEOL(rawInput, "linux")
```

## ServerObject: G3FileUploader Function: BlockExtension(ext)
**Keywords:** upload, security, extension, block, filter, restriction
**Description:** Adds a specific file extension to the blacklist.
Input: `ext` (String) - The extension to block (e.g., ".exe", "php").
Return: `Empty`.
**Observations:** Extensions are case-insensitive and the leading dot is optional. Blocking ".asp" and ".exe" is highly recommended for security.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3FileUploader")
ax.BlockExtension ".exe"
ax.BlockExtension "msi"
```

## ServerObject: G3FileUploader Function: AllowExtension(ext)
**Keywords:** upload, security, extension, allow, whitelist, filter
**Description:** Adds a specific file extension to the whitelist.
Input: `ext` (String) - The extension to allow (e.g., ".jpg", "pdf").
Return: `Empty`.
**Observations:** Only effective if `SetUseAllowedOnly(True)` is called.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3FileUploader")
ax.AllowExtension ".png"
ax.AllowExtension ".jpg"
```

## ServerObject: G3FileUploader Function: SetUseAllowedOnly(enabled)
**Keywords:** upload, security, extension, whitelist, mode, filter
**Description:** Toggles strict whitelist mode for file uploads.
Input: `enabled` (Boolean) - If True, only extensions in the `AllowedExtensions` list will be accepted.
Return: `Empty`.
**Observations:** When enabled, the `BlockedExtensions` list is still checked but the `AllowedExtensions` list becomes the primary filter.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3FileUploader")
ax.AllowExtensions "jpg,png,gif"
ax.SetUseAllowedOnly True
```

## ServerObject: G3FileUploader Function: Process(fieldName, targetDir, newFileName)
**Keywords:** upload, save, file, multipart, form, processing, filesystem
**Description:** Processes a single file upload from a multipart/form-data request.
Input: `fieldName` (String) - The name of the HTML file input field.
Input: `targetDir` (String) - The virtual directory where the file should be saved.
Input: `newFileName` (String) - Optional. A custom name for the saved file.
Return: `Dictionary` (VTNativeObject) - A dictionary containing result metadata.
**Observations:** The returned dictionary includes: `IsSuccess`, `OriginalFileName`, `NewFileName`, `Size`, `MimeType`, `Extension`, `FinalPath`, `RelativePath`, and `ErrorMessage`. If `newFileName` is omitted, a unique name is generated unless `PreserveOriginalName` is set to True.
**Syntax:**
```vbscript
Dim ax, result
Set ax = Server.CreateObject("G3FileUploader")
ax.maxfilesize = 5 * 1024 * 1024 ' 5MB
Set result = ax.Process("fileInput", "/uploads/images", "")
If result("IsSuccess") Then
    Response.Write "File saved as: " & result("RelativePath")
Else
    Response.Write "Error: " & result("ErrorMessage")
End If
```

## ServerObject: G3FileUploader Function: ProcessAll(targetDir)
**Keywords:** upload, save, batch, multiple, files, multipart
**Description:** Processes all files uploaded in the current request.
Input: `targetDir` (String) - The virtual directory where files should be saved.
Return: `Array` (VTArray) - An array of Dictionary objects, one for each processed file.
**Observations:** Extremely efficient for multi-file upload forms. Each element in the array follows the same schema as the `Process` return value.
**Syntax:**
```vbscript
Dim ax, results, i
Set ax = Server.CreateObject("G3FileUploader")
results = ax.ProcessAll("/uploads/batch")
For i = 0 To UBound(results)
    If results(i)("IsSuccess") Then
        Response.Write "Saved: " & results(i)("OriginalFileName") & "<br>"
    End If
Next
```

## ServerObject: G3FileUploader Function: GetFileInfo(fieldName)
**Keywords:** upload, file, info, metadata, check, validation
**Description:** Retrieves metadata about an uploaded file without saving it to disk.
Input: `fieldName` (String) - The name of the HTML file input field.
Return: `Dictionary` (VTNativeObject) - Metadata including `OriginalFileName`, `Size`, `MimeType`, `Extension`, `IsValid`, and `ExceedsMaxSize`.
**Observations:** Useful for pre-validation before calling `Process`.
**Syntax:**
```vbscript
Dim ax, info
Set ax = Server.CreateObject("G3FileUploader")
Set info = ax.GetFileInfo("fileInput")
If info("IsValid") Then
    Response.Write "File is acceptable"
End If
```

## ServerObject: G3Http Function: Fetch(url, method, body)
**Keywords:** http, request, get, post, api, fetch, json, remote, network, webservice
**Description:** Executes a remote HTTP request and returns the response.
Input: `url` (String) - The target URL.
Input: `method` (String) - Optional. HTTP method (GET, POST, PUT, DELETE, etc.). Default is "GET".
Input: `body` (String) - Optional. The request payload for POST/PUT.
Return: `Variant` - Returns a `String` if the response is plain text, or a `Dictionary`/`Array` if the response is `application/json`.
**Observations:** This is a high-level wrapper for Go's `http.Client`. It automatically parses JSON responses into native VBScript-accessible structures using the internal G3JSON parser. Timeout is fixed at 10 seconds. If the request fails, it returns `Empty`.
**Syntax:**
```vbscript
Dim ax, response
Set ax = Server.CreateObject("G3Http")
' Simple GET request
response = ax.Fetch("https://api.example.com/data")
Response.Write response

' POST request with JSON body
Dim payload
payload = "{""name"":""John""}"
Set response = ax.Fetch("https://api.example.com/users", "POST", payload)
If Not IsEmpty(response) Then
    Response.Write "New User ID: " & response("id")
End If
```

## ServerObject: G3Axon.Functions Function: AxEngineName()
**Keywords:** engine, name, identity, system, vm, identification, versioning, axonvm, runtime
**Description:** Returns the formal identification string of the AxonASP Virtual Machine engine. This is a constant string used to verify that the script is running under the AxonASP environment. Expected return type is a `String`.
**Observations:** This function is read-only and returns "AxonASP". It is used for environment detection and ensuring compatibility with Axon-specific features.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Engine: " & ax.AxEngineName()
```

## ServerObject: G3Axon.Functions Function: AxVersion()
**Keywords:** version, build, engine, release, metadata, versioning, runtime
**Description:** Returns the current version string of the AxonVM engine. It identifies the specific build and release level of the Go-based backend. Expected return type is a `String`.
**Observations:** The version follows the semantic versioning pattern (e.g., "1.1.3"). Useful for debugging and conditional feature usage based on the engine version.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "AxonASP Version: " & ax.AxVersion()
```

## ServerObject: G3Axon.Functions Function: AxGetEnv(name)
**Keywords:** environment, variable, os, settings, system, env, getenv, configuration
**Description:** Retrieves the value of an operating system environment variable.
Input: `name` (String) - The case-sensitive (on Linux) or case-insensitive (on Windows) name of the environment variable to retrieve.
Return: `String` - The value of the variable, or an empty string if the variable does not exist.
**Observations:** This provides direct access to the host environment. On Windows, environment variable names are case-insensitive, while on Linux/Unix systems, they are case-sensitive.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "System Path: " & ax.AxGetEnv("PATH")
```

## ServerObject: G3Axon.Functions Function: AxShutdownAxonASPServer()
**Keywords:** shutdown, exit, stop, server, maintenance, process, termination
**Description:** Triggers a controlled shutdown of the AxonASP server process.
Return: `Boolean` - Returns `True` if the shutdown signal was sent, `False` if the function is disabled in the configuration.
**Observations:** For security reasons, this function is disabled by default. It must be explicitly enabled in `config/axonasp.toml` under `axfunctions.enable_axservershutdown_function`. When called and enabled, the process will terminate immediately with a specific exit code.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
' This will terminate the server process if enabled
ax.AxShutdownAxonASPServer()
```

## ServerObject: G3Axon.Functions Function: AxChangeDir(path)
**Keywords:** directory, folder, path, cd, workdir, change, filesystem, context
**Description:** Changes the current working directory of the AxonASP process to the specified path.
Input: `path` (String) - The destination directory path.
Return: `Boolean` - Returns `True` on success, `False` if the path is invalid or inaccessible.
**Observations:** Changing the working directory affects all subsequent relative path operations (like `FileSystemObject` calls) for the entire process. Use with extreme caution in a multi-threaded web server environment.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
If ax.AxChangeDir("C:\Data\Logs") Then
    Response.Write "Directory changed successfully."
End If
```

## ServerObject: G3Axon.Functions Function: AxCurrentDir()
**Keywords:** directory, pwd, path, workdir, folder, current, filesystem, location
**Description:** Retrieves the absolute path of the current working directory of the AxonASP process.
Return: `String` - The absolute path of the current directory.
**Observations:** Equivalent to the `pwd` command in Unix or `cd` in Windows CMD. Useful for resolving paths relative to the process's current execution context.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Active Directory: " & ax.AxCurrentDir()
```

## ServerObject: G3Axon.Functions Function: AxHostnameValue()
**Keywords:** host, computer, name, network, identity, machine, hostname, dns
**Description:** Returns the network hostname of the machine where the AxonASP server is running.
Return: `String` - The system hostname as reported by the OS.
**Observations:** This is derived from the Go `os.Hostname()` call. It is useful for identifying specific nodes in a load-balanced environment.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Server Node: " & ax.AxHostnameValue()
```

## ServerObject: G3Axon.Functions Function: AxClearEnvironment()
**Keywords:** environment, env, clear, reset, security, variables, system
**Description:** Wipes all environment variables from the current process.
Return: `Boolean` - Always returns `True` after execution.
**Observations:** This is a high-impact operation that removes all inherited and set environment variables. Use only in specialized security scenarios or before spawning child processes with a clean environment.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
ax.AxClearEnvironment()
```

## ServerObject: G3Axon.Functions Function: AxEnvironmentList()
**Keywords:** environment, env, variables, list, array, enumeration, system, dump
**Description:** Returns an array containing all environment variables currently defined in the process.
Return: `Array` (VTArray) - A zero-based array where each element is a string in the "KEY=VALUE" format.
**Observations:** Provides a complete snapshot of the environment. The order of variables is determined by the underlying operating system.
**Syntax:**
```vbscript
Dim ax, envs, i
Set ax = Server.CreateObject("G3Axon.Functions")
envs = ax.AxEnvironmentList()
For i = 0 To UBound(envs)
    Response.Write envs(i) & "<br>"
Next
```

## ServerObject: G3Axon.Functions Function: AxEnvironmentValue(name [, default])
**Keywords:** environment, variable, env, lookup, settings, default, system
**Description:** Retrieves the value of a specific environment variable, with an optional fallback default.
Input: `name` (String) - The name of the environment variable.
Input: `default` (Variant) - Optional. The value to return if the variable is not found.
Return: `Variant` - The string value of the environment variable or the provided default.
**Observations:** Unlike `AxGetEnv`, this function allows providing a default value directly in the call, simplifying logic for missing configurations.
**Syntax:**
```vbscript
Dim ax, port
Set ax = Server.CreateObject("G3Axon.Functions")
port = ax.AxEnvironmentValue("APP_PORT", "8080")
Response.Write "Running on port: " & port
```

## ServerObject: G3Axon.Functions Function: AxProcessId()
**Keywords:** pid, process, identity, system, os, task, monitoring
**Description:** Returns the unique numeric Process ID (PID) assigned by the operating system to the current AxonASP instance.
Return: `Integer` (int64).
**Observations:** Useful for generating unique filenames, logging for process monitoring, or diagnostic purposes.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Process ID: " & ax.AxProcessId()
```

## ServerObject: G3Axon.Functions Function: AxEffectiveUserId()
**Keywords:** uid, user, identity, effective, permissions, system, os, security
**Description:** Returns the effective numeric user ID of the caller.
Return: `Integer` (int64).
**Observations:** On Windows systems, this function returns -1 as UIDs are not a native concept. On Linux/Unix, it returns the `euid` of the process.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
If ax.AxEffectiveUserId() = 0 Then
    Response.Write "Running as root/superuser"
End If
```

## ServerObject: G3Axon.Functions Function: AxDirSeparator()
**Keywords:** path, separator, slash, backslash, os, filesystem, cross-platform
**Description:** Returns the character used by the operating system to separate directory components in a file path.
Return: `String` - Returns "\" on Windows and "/" on Linux/Unix systems.
**Observations:** Essential for building portable file paths that work across different hosting environments.
**Syntax:**
```vbscript
Dim ax, sep
Set ax = Server.CreateObject("G3Axon.Functions")
sep = ax.AxDirSeparator()
Response.Write "Relative Path: docs" & sep & "manual.pdf"
```

## ServerObject: G3Axon.Functions Function: AxPathListSeparator()
**Keywords:** path, separator, list, classpath, os, environment
**Description:** Returns the character used by the operating system to separate multiple paths in a list (like the PATH environment variable).
Return: `String` - Returns ";" on Windows and ":" on Linux/Unix.
**Observations:** Useful when parsing or constructing environment variables that contain multiple file paths.
**Syntax:**
```vbscript
Dim ax, sep
Set ax = Server.CreateObject("G3Axon.Functions")
sep = ax.AxPathListSeparator()
Response.Write "The PATH list separator is: " & sep
```

## ServerObject: G3Axon.Functions Function: AxIntegerSizeBytes()
**Keywords:** integer, size, bytes, architecture, memory, limit
**Description:** Returns the size in bytes of a standard integer on the current platform.
Return: `Integer`.
**Observations:** Returns 4 on 32-bit systems and 8 on 64-bit systems. Helps determine the memory footprint and limits of integer values.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Integer size: " & ax.AxIntegerSizeBytes() & " bytes"
```

## ServerObject: G3Axon.Functions Function: AxPlatformBits()
**Keywords:** platform, bits, architecture, 32-bit, 64-bit, cpu, system
**Description:** Returns the architecture bit-size of the current platform.
Return: `Integer` - Typically 32 or 64.
**Observations:** Useful for system-level diagnostics and determining the capabilities of the underlying hardware and OS.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Running on a " & ax.AxPlatformBits() & "-bit system"
```

## ServerObject: G3Axon.Functions Function: AxExecutablePath()
**Keywords:** executable, path, binary, binary location, system, location
**Description:** Returns the absolute path of the AxonASP executable currently running the script.
Return: `String` - The full path to the .exe (Windows) or binary (Linux).
**Observations:** Useful for finding configuration files or resources located in the same directory as the server executable.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Executable location: " & ax.AxExecutablePath()
```

## ServerObject: G3Axon.Functions Function: AxExecute(command)
**Keywords:** exec, shell, command, system, run, output, cmd, bash, execute
**Description:** Executes an external shell command and captures the output.
Input: `command` (String) - The full command line to execute.
Return: `String` - The combined output from both standard output (stdout) and standard error (stderr), with trailing newlines removed.
**Observations:** Uses `cmd.exe /c` on Windows and `sh -c` on Unix-like systems. This is a blocking call; the script will wait until the command completes. Returns `False` (Boolean) if the command is empty.
**Syntax:**
```vbscript
Dim ax, output
Set ax = Server.CreateObject("G3Axon.Functions")
output = ax.AxExecute("dir C:\") ' Windows example
Response.Write "<pre>" & output & "</pre>"
```

## ServerObject: G3Axon.Functions Function: AxSystemInfo([mode])
**Keywords:** system, info, os, version, architecture, hostname, platform, metadata
**Description:** Retrieves detailed system or runtime information based on the requested mode.
Input: `mode` (String) - Optional. "s" (OS name), "n" (Hostname), "v" (Go runtime version), "m" (Architecture). If omitted or "a", returns a combined string of all info.
Return: `String`.
**Observations:** High-level summary of the hosting environment. Case-insensitive mode flags.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "OS: " & ax.AxSystemInfo("s") & "<br>"
Response.Write "Arch: " & ax.AxSystemInfo("m") & "<br>"
Response.Write "Full: " & ax.AxSystemInfo("a")
```

## ServerObject: G3Axon.Functions Function: AxCurrentUser()
**Keywords:** user, identity, username, current, session, security, owner
**Description:** Returns the username of the user currently running the AxonASP process.
Return: `String`.
**Observations:** On Windows, it attempts to get the `USERNAME` environment variable if the user lookup fails. On Unix, it uses the `USER` environment variable as a fallback.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Process Owner: " & ax.AxCurrentUser()
```

## ServerObject: G3Axon.Functions Function: Axw(string)
**Keywords:** html, escape, security, xss, entities, sanitize, write, output
**Description:** Escapes special HTML characters and writes the result directly to the HTTP response buffer.
Input: `string` (String) - The content to escape and write.
Return: `Empty`.
**Observations:** This is a security-focused version of `Response.Write`. It automatically converts characters like `<`, `>`, `&`, `"`, and `'` into their HTML entity equivalents to prevent Cross-Site Scripting (XSS). Note: Also responds to `document.write` and `documentwrite` method names for compatibility.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
' This will display as literal text in the browser, not as an alert
ax.Axw("<script>alert('XSS')</script>")
```

## ServerObject: G3Axon.Functions Function: AxMax(args...)
**Keywords:** math, max, maximum, comparison, largest, numbers
**Description:** Returns the largest value among the provided numeric arguments.
Input: `args...` (Variadic) - One or more numeric values (Integer or Double).
Return: `Double`.
**Observations:** If no arguments are provided, returns 0. If multiple types are mixed, all values are coerced to `Double` for comparison.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Max: " & ax.AxMax(10, 55.5, 22, 104) ' Returns 104
```

## ServerObject: G3Axon.Functions Function: AxMin(args...)
**Keywords:** math, min, minimum, comparison, smallest, numbers
**Description:** Returns the smallest value among the provided numeric arguments.
Input: `args...` (Variadic) - One or more numeric values (Integer or Double).
Return: `Double`.
**Observations:** If no arguments are provided, returns 0. All values are coerced to `Double` during comparison.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Min: " & ax.AxMin(10, 55.5, 5, 104) ' Returns 5
```

## ServerObject: G3Axon.Functions Function: AxIntegerMax()
**Keywords:** math, integer, max, limit, boundary, 64-bit
**Description:** Returns the maximum possible value for a signed integer on the current platform.
Return: `Integer` (int64).
**Observations:** On 64-bit systems, this returns 9,223,372,036,854,775,807.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Max Int: " & ax.AxIntegerMax()
```

## ServerObject: G3Axon.Functions Function: AxIntegerMin()
**Keywords:** math, integer, min, limit, boundary, 64-bit
**Description:** Returns the minimum possible value for a signed integer on the current platform.
Return: `Integer` (int64).
**Observations:** On 64-bit systems, this returns -9,223,372,036,854,775,808.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Min Int: " & ax.AxIntegerMin()
```

## ServerObject: G3Axon.Functions Function: AxCeil(number)
**Keywords:** math, ceil, ceiling, round, rounding, upward
**Description:** Returns the smallest integer greater than or equal to the specified number.
Input: `number` (Double).
Return: `Double`.
**Observations:** Coerces input to `Double`. Equivalent to `math.Ceil` in Go.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxCeil(4.2) ' Returns 5
```

## ServerObject: G3Axon.Functions Function: AxFloor(number)
**Keywords:** math, floor, round, rounding, downward
**Description:** Returns the largest integer less than or equal to the specified number.
Input: `number` (Double).
Return: `Double`.
**Observations:** Coerces input to `Double`. Equivalent to `math.Floor` in Go.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxFloor(4.8) ' Returns 4
```

## ServerObject: G3Axon.Functions Function: AxRand([min, max])
**Keywords:** random, rand, number, integer, dice, rng, stochastic
**Description:** Generates a pseudo-random integer.
Input: `min` (Integer) - Optional. The lower bound.
Input: `max` (Integer) - Optional. The upper bound.
Return: `Integer`.
**Observations:** 
- If no arguments: returns a random positive integer.
- If one argument `max`: returns a random integer between 0 and `max` (inclusive).
- If two arguments `min, max`: returns a random integer between `min` and `max` (inclusive).
Automatically swaps bounds if `min > max`.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
' Random number between 1 and 100
Response.Write "Roll: " & ax.AxRand(1, 100)
```

## ServerObject: G3Axon.Functions Function: AxNumberFormat(num [, decimals, decPoint, thousandsSep])
**Keywords:** math, format, number, currency, precision, string manipulation
**Description:** Formats a number with grouped thousands and custom decimal separators.
Input: `num` (Double) - The number to format.
Input: `decimals` (Integer) - Optional. Number of decimal points (default: 2).
Input: `decPoint` (String) - Optional. Decimal separator (default: ".").
Input: `thousandsSep` (String) - Optional. Thousands separator (default: ",").
Return: `String`.
**Observations:** Replicates the PHP `number_format` behavior. Ideal for localized financial display.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
' Formats as "1.234.567,89"
Response.Write ax.AxNumberFormat(1234567.8912, 2, ",", ".")
```

## ServerObject: G3Axon.Functions Function: AxPi()
**Keywords:** math, pi, constant, circle, geometry
**Description:** Returns the mathematical constant Pi (π).
Return: `Double`.
**Observations:** Highly precise value (3.141592653589793...).
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Pi is roughly: " & ax.AxPi()
```

## ServerObject: G3Axon.Functions Function: AxSmallestFloatValue()
**Keywords:** math, float, epsilon, limit, precision, double
**Description:** Returns the smallest positive non-zero value that can be represented by a `Double`.
Return: `Double`.
**Observations:** Equivalent to `math.SmallestNonzeroFloat64` in Go. Useful for epsilon comparisons.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Smallest Float: " & ax.AxSmallestFloatValue()
```

## ServerObject: G3Axon.Functions Function: AxGetConfig()
**Keywords:** config, settings, configuration, toml, viper
**Description:** Retrieves a configuration value from the AxonASP configuration file (`axonasp.toml`).
Input: `key` (String) - The configuration key to retrieve.
Return: `Variant` - The value of the configuration key, or `Empty` if not found.
**Observations:** Useful for accessing runtime configuration without hard-coding values. If an env file is used with the same config name, it will use the env variable. 
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Config Value: " & ax.AxGetConfig("some.config.key")
```


## ServerObject: G3Axon.Functions Function: AxGetConfigKeys()
**Keywords:** config, settings, configuration, toml, viper
**Description:** Retrieves all configuration keys from the AxonASP configuration file (`axonasp.toml`).
Return: `Array` (VTArray) - An array of all configuration keys.
**Observations:** Useful for accessing runtime configuration without hard-coding values.
**Syntax:**

```vbscript
Dim ax, keys
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")
keys = ax.axgetconfigkeys()
```

## ServerObject: G3Axon.Functions Function: AxGetDefaultCSS()
**Keywords:** css, theme, default style, design, appearance
**Description:** Returns a string containing the path to default CSS stylesheet used by AxonASP for basic HTML.
**Observations:** Useful for referencing the AxonASP default stylesheet path at runtime without hard-coding the location in every ASP page.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Default CSS: " & ax.AxGetDefaultCSS()
```

## ServerObject: G3Axon.Functions Function: AxFloatPrecisionDigits()
**Keywords:** math, float, precision, digits, double, limit
**Description:** Returns the number of decimal digits of precision for a `Double`.
Return: `Integer` - Returns 15.
**Observations:** Constant value representing standard IEEE-754 double precision capability.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Double precision: " & ax.AxFloatPrecisionDigits() & " digits"
```

## ServerObject: G3Axon.Functions Function: AxCount(array)
**Keywords:** count, length, size, array, elements, collection, ubound
**Description:** Returns the total number of elements in a zero-based or multi-dimensional array.
Input: `array` (VTArray) - The array to measure.
Return: `Integer`.
**Observations:** Returns 0 if the input is not an array. Faster than calculating `UBound(arr) - LBound(arr) + 1` manually.
**Syntax:**
```vbscript
Dim ax, myArr
Set ax = Server.CreateObject("G3Axon.Functions")
myArr = Array("A", "B", "C")
Response.Write "Array has " & ax.AxCount(myArr) & " items"
```

## ServerObject: G3Axon.Functions Function: AxExplode(delimiter, string [, limit])
**Keywords:** split, explode, string, array, delimiter, tokens, parsing
**Description:** Splits a string into an array of strings based on a delimiter.
Input: `delimiter` (String) - The boundary string.
Input: `string` (String) - The input string to split.
Input: `limit` (Integer) - Optional. Max number of elements in the resulting array.
Return: `Array` (VTArray).
**Observations:** If delimiter is empty, the string is split into individual characters. If `limit` is specified, the resulting array will contain at most `limit` elements.
**Syntax:**
```vbscript
Dim ax, parts, i
Set ax = Server.CreateObject("G3Axon.Functions")
parts = ax.AxExplode("|", "apple|orange|banana|grape", 3)
For i = 0 To UBound(parts)
    Response.Write i & ": " & parts(i) & "<br>"
Next
' Result: 0: apple, 1: orange, 2: banana
```

## ServerObject: G3Axon.Functions Function: AxImplode(glue, array)
**Keywords:** join, implode, glue, string, array, concatenation, flattening
**Description:** Joins all elements of an array into a single string, with a separator between each element.
Input: `glue` (String) - The string to place between elements.
Input: `array` (VTArray) - The array of elements to join.
Return: `String`.
**Observations:** All array elements are automatically converted to their string representations before joining.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxImplode(" -> ", Array("Step 1", "Step 2", "Step 3"))
```

## ServerObject: G3Axon.Functions Function: AxArrayReverse(array)
**Keywords:** array, reverse, flip, order, sorting
**Description:** Returns a new array with the order of elements reversed.
Input: `array` (VTArray) - The source array.
Return: `Array` (VTArray).
**Observations:** If input is not an array, it returns the input value as-is. Performs a shallow copy of the elements into the new order.
**Syntax:**
```vbscript
Dim ax, rev, i
Set ax = Server.CreateObject("G3Axon.Functions")
rev = ax.AxArrayReverse(Array(1, 2, 3, 4, 5))
' rev will be Array(5, 4, 3, 2, 1)
```

## ServerObject: G3Axon.Functions Function: AxRange(start, end [, step])
**Keywords:** array, range, sequence, numbers, loop, generator
**Description:** Creates an array containing a sequence of integers.
Input: `start` (Integer) - The beginning of the sequence.
Input: `end` (Integer) - The end of the sequence.
Input: `step` (Integer) - Optional. The increment/decrement between values (default: 1).
Return: `Array` (VTArray).
**Observations:** Supports both increasing and decreasing sequences based on the `step` value. If `step` is 0, it defaults to 1.
**Syntax:**
```vbscript
Dim ax, seq
Set ax = Server.CreateObject("G3Axon.Functions")
seq = ax.AxRange(10, 50, 10)
' seq will be Array(10, 20, 30, 40, 50)
```

## ServerObject: G3Axon.Functions Function: AxStringReplace(search, replace, subject)
**Keywords:** string, replace, substitution, swap, manipulation
**Description:** Replaces all occurrences of a search string with a replacement string within a subject string.
Input: `search` (String) - The string to find.
Input: `replace` (String) - The replacement string.
Input: `subject` (String) - The text to process.
Return: `String`.
**Observations:** Performs a global, case-sensitive replacement. Equivalent to Go's `strings.ReplaceAll`.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxStringReplace("red", "blue", "The red car and red bike.")
' Output: "The blue car and blue bike."
```

## ServerObject: G3Axon.Functions Function: AxPad(str, length [, padString, padType])
**Keywords:** string, pad, padding, formatting, alignment
**Description:** Pads a string to a certain length with another string.
Input: `str` (String) - The input string.
Input: `length` (Integer) - The target length.
Input: `padString` (String) - Optional. The string to use for padding (default: space).
Input: `padType` (Integer) - Optional. 0 (Left), 1 (Right), 2 (Both). Default: 1 (Right).
Return: `String`.
**Observations:** If the input string is already longer than `length`, it is returned unchanged.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
' Pads with zeros on the left: "00042"
Response.Write ax.AxPad("42", 5, "0", 0)
```

## ServerObject: G3Axon.Functions Function: AxRepeat(str, times)
**Keywords:** string, repeat, duplicate, multiply, generation
**Description:** Returns a string consisting of the input string repeated a specified number of times.
Input: `str` (String) - The string to repeat.
Input: `times` (Integer) - Number of repetitions.
Return: `String`.
**Observations:** If `times` is less than 1, returns an empty string. Highly efficient for generating separator lines or repeated patterns.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxRepeat("-", 50) ' Prints 50 dashes
```

## ServerObject: G3Axon.Functions Function: AxUcFirst(string)
**Keywords:** string, uppercase, capitalize, format, casing
**Description:** Converts the first character of a string to uppercase.
Input: `string` (String).
Return: `String`.
**Observations:** Works correctly with Unicode characters (runes). Only affects the first character; the rest of the string is left unchanged.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxUcFirst("hello axon") ' Returns "Hello axon"
```

## ServerObject: G3Axon.Functions Function: AxWordCount(string [, format])
**Keywords:** string, word count, statistics, analysis, array, parsing
**Description:** Counts the number of words in a string, or returns an array containing the words.
Input: `string` (String).
Input: `format` (Integer) - Optional. 0 (returns count), 1 (returns array of words). Default: 0.
Return: `Variant` - Either an Integer (count) or an Array (words).
**Observations:** Uses whitespace as the word boundary.
**Syntax:**
```vbscript
Dim ax, words, count
Set ax = Server.CreateObject("G3Axon.Functions")
count = ax.AxWordCount("One two three")
words = ax.AxWordCount("One two three", 1)
Response.Write "Count: " & count & ", First word: " & words(0)
```

## ServerObject: G3Axon.Functions Function: AxNl2Br(string)
**Keywords:** string, newline, br, html, format, conversion, line break
**Description:** Replaces all newline characters (`\r\n`, `\n`, `\r`) in a string with HTML `<br>` tags.
Input: `string` (String).
Return: `String`.
**Observations:** Useful for displaying plain text from textareas or files in an HTML page while preserving line breaks.
**Syntax:**
```vbscript
Dim ax, text
Set ax = Server.CreateObject("G3Axon.Functions")
text = "Line 1" & vbCrLf & "Line 2"
Response.Write ax.AxNl2Br(text)
```

## ServerObject: G3Axon.Functions Function: AxTrim(string [, chars])
**Keywords:** string, trim, strip, whitespace, cleanup, sanitize
**Description:** Removes whitespace or other specified characters from the beginning and end of a string.
Input: `string` (String).
Input: `chars` (String) - Optional. A string containing characters to remove (default: standard whitespace).
Return: `String`.
**Observations:** Standard whitespace includes space, tab, newline, carriage return, vertical tab, and form feed.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "[" & ax.AxTrim("   data   ") & "]" ' Returns "[data]"
Response.Write ax.AxTrim("!!hello!!", "!") ' Returns "hello"
```

## ServerObject: G3Axon.Functions Function: AxStringGetCSV(string [, delimiter])
**Keywords:** string, csv, parse, array, delimiter, excel, data
**Description:** Parses a single line of CSV-formatted text and returns an array of fields.
Input: `string` (String) - The CSV line to parse.
Input: `delimiter` (String) - Optional. The field separator character (default: ",").
Return: `Array` (VTArray).
**Observations:** Correcty handles quoted fields containing the delimiter or escaped quotes. Based on Go's `encoding/csv` package.
**Syntax:**
```vbscript
Dim ax, fields
Set ax = Server.CreateObject("G3Axon.Functions")
fields = ax.AxStringGetCSV("101,""Lucas, G."",True")
Response.Write "Name: " & fields(1) ' Returns "Lucas, G."
```

## ServerObject: G3Axon.Functions Function: AxMd5(string)
**Keywords:** md5, hash, crypto, security, digest, string, fingerprint
**Description:** Calculates the MD5 (Message-Digest algorithm 5) hash of a string.
Input: `string` (String).
Return: `String` - A 32-character hexadecimal string.
**Observations:** Fast hashing algorithm. While MD5 is considered cryptographically broken for high-security applications, it remains widely used for data integrity checks and non-sensitive fingerprints.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Hash: " & ax.AxMd5("axonasp")
```

## ServerObject: G3Axon.Functions Function: AxSha1(string)
**Keywords:** sha1, hash, crypto, security, digest, string, fingerprint
**Description:** Calculates the SHA-1 (Secure Hash Algorithm 1) hash of a string.
Input: `string` (String).
Return: `String` - A 40-character hexadecimal string.
**Observations:** Similar to MD5 but produces a longer digest. Should not be used for sensitive password storage.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "SHA1: " & ax.AxSha1("axonasp")
```

## ServerObject: G3Axon.Functions Function: AxHash(algo, string)
**Keywords:** hash, crypto, sha256, sha1, md5, security, digest
**Description:** Calculates a cryptographic hash using the specified algorithm.
Input: `algo` (String) - "sha256", "sha1", or "md5".
Input: `string` (String) - The data to hash.
Return: `String` - Hexadecimal digest string.
**Observations:** Centralized function for multiple hashing algorithms. SHA-256 is the recommended choice for modern security requirements.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "SHA256: " & ax.AxHash("sha256", "axonasp")
```

## ServerObject: G3Axon.Functions Function: AxBase64Encode(string)
**Keywords:** base64, encode, encoding, binary, string, transport
**Description:** Encodes a string into Base64 format.
Input: `string` (String).
Return: `String`.
**Observations:** Useful for transmitting binary-like data over text-based protocols (like HTML or JSON).
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxBase64Encode("Hello") ' Returns "SGVsbG8="
```

## ServerObject: G3Axon.Functions Function: AxBase64Decode(string)
**Keywords:** base64, decode, decoding, binary, string
**Description:** Decodes a Base64 encoded string back to its original plain text.
Input: `string` (String) - Base64 encoded string.
Return: `String`.
**Observations:** Returns an empty string if the input is not valid Base64.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxBase64Decode("SGVsbG8=") ' Returns "Hello"
```

## ServerObject: G3Axon.Functions Function: AxUrlDecode(string)
**Keywords:** url, decode, decoding, query, escape, percent-encoding
**Description:** Decodes a URL-encoded string (percent-encoding).
Input: `string` (String).
Return: `String`.
**Observations:** Converts `%XX` sequences into their character equivalents.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxUrlDecode("Hello%20World%21") ' Returns "Hello World!"
```

## ServerObject: G3Axon.Functions Function: AxRawUrlDecode(string)
**Keywords:** url, decode, raw, escape, percent-encoding
**Description:** Decodes a URL-encoded string, similar to `AxUrlDecode` but specifically optimized for raw query components.
Input: `string` (String).
Return: `String`.
**Observations:** Explicitly handles `+` characters as spaces before decoding.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxRawUrlDecode("Hello+World") ' Returns "Hello World"
```

## ServerObject: G3Axon.Functions Function: AxRgbToHex(r, g, b)
**Keywords:** color, rgb, hex, conversion, formatting, ui, graphics
**Description:** Converts Red, Green, and Blue integer components into a hexadecimal color string.
Input: `r, g, b` (Integer) - 0 to 255.
Return: `String` - Format "#RRGGBB".
**Observations:** Automatically clamps or masks values to the 0-255 range.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Hex color: " & ax.AxRgbToHex(255, 128, 0) ' Returns "#FF8000"
```

## ServerObject: G3Axon.Functions Function: AxHtmlSpecialChars(string)
**Keywords:** html, escape, security, xss, entities, sanitize, formatting
**Description:** Escapes special HTML characters to prevent cross-site scripting (XSS).
Input: `string` (String).
Return: `String`.
**Observations:** Converts `<`, `>`, `&`, `"`, and `'`. This should be used whenever displaying user-provided data in an HTML context.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxHtmlSpecialChars("<b onclick='bad()'>text</b>")
' Output: &lt;b onclick=&#39;bad()&#39;&gt;text&lt;/b&gt;
```

## ServerObject: G3Axon.Functions Function: AxStripTags(string)
**Keywords:** html, strip, tags, cleanup, sanitize, text, plain text
**Description:** Removes all HTML and PHP tags from a string.
Input: `string` (String).
Return: `String`.
**Observations:** Uses a regular expression to identify and remove anything within `< >`. Useful for extracting plain text from HTML content.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxStripTags("<p>Hello <b>World</b></p>") ' Returns "Hello World"
```

## ServerObject: G3Axon.Functions Function: AxFilterValidateIP(ip)
**Keywords:** ip, validation, network, filter, check, security, ipv4, ipv6
**Description:** Validates whether a string is a valid IP address (IPv4 or IPv6).
Input: `ip` (String).
Return: `Boolean`.
**Observations:** Uses Go's `net.ParseIP` for high-reliability validation.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
If ax.AxFilterValidateIP("192.168.1.1") Then
    Response.Write "Valid IP"
End If
```

## ServerObject: G3Axon.Functions Function: AxFilterValidateEmail(email)
**Keywords:** email, validation, filter, check, security, regex
**Description:** Validates whether a string is a valid email address format.
Input: `email` (String).
Return: `Boolean`.
**Observations:** Uses RFC 5322 compliant parsing via Go's `net/mail` package.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
If ax.AxFilterValidateEmail("user@example.com") Then
    Response.Write "Valid Email"
End If
```

## ServerObject: G3Axon.Functions Function: AxIsInt(value)
**Keywords:** type, check, validation, integer, numeric
**Description:** Checks if the provided value is of the internal `VTInteger` type.
Input: `value` (Variant).
Return: `Boolean`.
**Observations:** Strict type check. A string containing a number will return `False`. Use `IsNumeric` for loose checking.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxIsInt(42) ' True
Response.Write ax.AxIsInt("42") ' False
```

## ServerObject: G3Axon.Functions Function: AxIsFloat(value)
**Keywords:** type, check, validation, float, double, numeric
**Description:** Checks if the provided value is of the internal `VTDouble` (floating point) type.
Input: `value` (Variant).
Return: `Boolean`.
**Observations:** Strict type check for floating point numbers.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxIsFloat(3.14) ' True
```

## ServerObject: G3Axon.Functions Function: AxCtypeAlpha(string)
**Keywords:** validation, alpha, letters, check, string
**Description:** Checks if a string contains only alphabetic characters (A-Z, a-z).
Input: `string` (String).
Return: `Boolean`.
**Observations:** Returns `False` if the string contains spaces, numbers, or punctuation.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxCtypeAlpha("Axon") ' True
Response.Write ax.AxCtypeAlpha("Axon 2") ' False
```

## ServerObject: G3Axon.Functions Function: AxCtypeAlnum(string)
**Keywords:** validation, alphanumeric, letters, numbers, check, string
**Description:** Checks if a string contains only alphanumeric characters (A-Z, a-z, 0-9).
Input: `string` (String).
Return: `Boolean`.
**Observations:** Useful for simple username or ID validation.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write ax.AxCtypeAlnum("Axon2026") ' True
```

## ServerObject: G3Axon.Functions Function: AxEmpty(value)
**Keywords:** validation, empty, null, zero, check, logic
**Description:** Determines whether a variable is "empty" based on a multi-type check.
Input: `value` (Variant).
Return: `Boolean`.
**Observations:** Returns `True` for: `VTEmpty`, `VTNull`, empty strings (""), integer 0, double 0.0, and boolean False.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
If ax.AxEmpty(Request.Form("optional_field")) Then
    Response.Write "Field was not provided or is zero."
End If
```

## ServerObject: G3Axon.Functions Function: AxIsSet(value)
**Keywords:** validation, isset, check, definition, initialization
**Description:** Checks if a variable has been initialized and is not `Empty` or `Null`.
Input: `value` (Variant).
Return: `Boolean`.
**Observations:** Inverse of a pure `VTEmpty`/`VTNull` check. Does not consider 0 or empty strings as "not set".
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
If ax.AxIsSet(Session("UserID")) Then
    Response.Write "User is logged in."
End If
```

## ServerObject: G3Axon.Functions Function: AxTime()
**Keywords:** time, unix, timestamp, now, epoch, clock
**Description:** Returns the current Unix timestamp (number of seconds elapsed since January 1, 1970 UTC).
Return: `Integer` (int64).
**Observations:** Respects the server's local time zone if configured in the VM, but typically returns UTC-relative epoch seconds.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Current Epoch: " & ax.AxTime()
```

## ServerObject: G3Axon.Functions Function: AxDate(format [, timestamp])
**Keywords:** date, format, time, parse, string, formatting, php-style
**Description:** Formats a local date/time using PHP-style formatting tokens.
Input: `format` (String) - Pattern string (Y: Year, m: Month, d: Day, H: Hour, i: Minute, s: Second, etc.).
Input: `timestamp` (Integer) - Optional. Unix timestamp to format (default: current time).
Return: `String`.
**Observations:** Powerful formatting engine. Escape characters with `\` to include them literally. Supports localized month and day names if `LCID` or `Locale` is set in the script.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
' "Today is Sunday, March 22, 2026"
Response.Write ax.AxDate("Today is l, F j, Y")
```

## ServerObject: G3Axon.Functions Function: AxLastModified()
**Keywords:** filesystem, file, modified, time, timestamp, metadata
**Description:** Returns the Unix timestamp of the last modification time of the currently executing ASP file.
Return: `Integer` (int64).
**Observations:** Automatically detects the file path of the script currently being interpreted by the VM.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Page updated: " & ax.AxDate("Y-m-d", ax.AxLastModified())
```

## ServerObject: G3Axon.Functions Function: AxGetRemoteFile(url)
**Keywords:** network, http, get, fetch, remote, file, download, requester
**Description:** Performs a synchronous HTTP GET request to the specified URL and returns the response body as a string.
Input: `url` (String) - Must start with `http://` or `https://`.
Return: `Variant` - The body content as a `String`, or `False` (Boolean) if the request fails or returns a non-200 status code.
**Observations:** Has a hardcoded 10-second timeout. Not intended for large file transfers; use for small APIs or config fetches.
**Syntax:**
```vbscript
Dim ax, json
Set ax = Server.CreateObject("G3Axon.Functions")
json = ax.AxGetRemoteFile("https://api.example.com/data.json")
If json <> False Then
    Response.Write "Data: " & json
End If
```

## ServerObject: G3Axon.Functions Function: AxGenerateGuid()
**Keywords:** guid, uuid, identity, unique, random, string, v4
**Description:** Generates a cryptographically secure Version 4 Universally Unique Identifier (UUID/GUID).
Return: `String` - Standard 36-character hyphenated format (e.g., "550e8400-e29b-41d4-a716-446655440000").
**Observations:** Guaranteed unique across time and space. Ideal for database primary keys, session identifiers, or temporary file names.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
Response.Write "Your ID: " & ax.AxGenerateGuid()
```

## ServerObject: G3Image Function: New(w, h)
**Keywords:** image, create, canvas, context, width, height, graphics
**Description:** Initializes a new drawing context with specific dimensions.
Input: `w` (Integer) - Width in pixels.
Input: `h` (Integer) - Height in pixels.
Return: `Boolean` - True if context created successfully.
**Observations:** Must be called before any drawing operation. Clears any previous state.
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
img.New 800, 600
```

## ServerObject: G3Image Function: Load(path)
**Keywords:** image, load, open, file, read, graphics
**Description:** Loads an image from the filesystem.
Input: `path` (String) - Path to the image file (PNG or JPG).
Return: `Boolean` - True if loaded successfully.
**Observations:** Resolves path relative to web root. Use `NewContextForImage` to draw on it.
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
If img.Load("/assets/bg.png") Then
    img.UseImage
End If
```

## ServerObject: G3Image Function: SavePNG(path)
**Keywords:** image, save, png, export, file, graphics
**Description:** Saves the current context as a PNG file.
Input: `path` (String) - Target file path.
Return: `Boolean` - True on success.
**Observations:** Automatically creates parent directories. Lossless format.
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
img.New 100, 100
img.SetHexColor "#FF0000"
img.DrawCircle 50, 50, 40
img.Fill
img.SavePNG "/output/circle.png"
```

## ServerObject: G3Image Function: SaveJPG(path, quality)
**Keywords:** image, save, jpg, jpeg, export, quality, compression
**Description:** Saves the current context as a JPEG file with configurable quality.
Input: `path` (String) - Target file path.
Input: `quality` (Integer) - Optional. 1-100 (default: 90).
Return: `Boolean` - True on success.
**Observations:** Lossy format, ideal for photographs.
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
' ... drawing ...
img.SaveJPG "/output/photo.jpg", 85
```

## ServerObject: G3Image Function: SetHexColor(hex)
**Keywords:** color, hex, style, drawing, context
**Description:** Sets the active drawing color using a hexadecimal string.
Input: `hex` (String) - Color hex (e.g., "#RRGGBB" or "#RRGGBBAA").
Return: `Empty`.
**Observations:** Supports alpha channel in hex (last 2 digits).
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
img.SetHexColor "#00FF00" ' Green
```

## ServerObject: G3Image Function: SetLineWidth(width)
**Keywords:** line, width, thickness, stroke, style
**Description:** Sets the thickness for subsequent stroke operations.
Input: `width` (Double) - Line thickness in pixels.
Return: `Empty`.
**Observations:** Affects `Stroke`, `DrawLine`, etc.
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
img.SetLineWidth 5.5
```

## ServerObject: G3Image Function: DrawLine(x1, y1, x2, y2)
**Keywords:** line, draw, path, graphics, coordinates
**Description:** Adds a line path from (x1, y1) to (x2, y2).
Input: `x1`, `y1` (Double) - Start coordinates.
Input: `x2`, `y2` (Double) - End coordinates.
Return: `Empty`.
**Observations:** Path must be finished with `Stroke`.
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
img.DrawLine 10, 10, 90, 90
img.Stroke
```

## ServerObject: G3Image Function: DrawRectangle(x, y, w, h)
**Keywords:** rectangle, box, draw, path, graphics
**Description:** Adds a rectangle path.
Input: `x`, `y` (Double) - Top-left corner.
Input: `w`, `h` (Double) - Width and height.
Return: `Empty`.
**Observations:** Path must be finished with `Stroke` or `Fill`.
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
img.DrawRectangle 20, 20, 150, 100
img.Fill
```

## ServerObject: G3Image Function: DrawCircle(x, y, r)
**Keywords:** circle, draw, path, graphics, radius
**Description:** Adds a circular path.
Input: `x`, `y` (Double) - Center coordinates.
Input: `r` (Double) - Radius in pixels.
Return: `Empty`.
**Observations:** Path must be finished with `Stroke` or `Fill`.
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
img.DrawCircle 100, 100, 50
img.Stroke
```

## ServerObject: G3Image Function: Stroke()
**Keywords:** stroke, draw, path, line, border
**Description:** Outlines the current path using the active color and line width.
Return: `Empty`.
**Observations:** Clears the current path after execution unless `StrokePreserve` is used.
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
img.DrawLine 0, 0, 100, 100
img.Stroke
```

## ServerObject: G3Image Function: Fill()
**Keywords:** fill, paint, background, path, area
**Description:** Fills the current path with the active color.
Return: `Empty`.
**Observations:** Clears the current path after execution unless `FillPreserve` is used.
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
img.DrawRectangle 0, 0, 50, 50
img.Fill
```

## ServerObject: G3Image Function: LoadFontFace(path, points)
**Keywords:** font, text, typography, load, graphics
**Description:** Loads a TrueType (TTF) font for text rendering.
Input: `path` (String) - Path to the .ttf file.
Input: `points` (Double) - Font size in points.
Return: `Boolean`.
**Observations:** Essential for `DrawString` operations.
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
img.LoadFontFace "/fonts/arial.ttf", 24
img.DrawString "Hello Axon", 50, 50
```

## ServerObject: G3Image Function: DrawString(s, x, y)
**Keywords:** text, draw, string, typography, print
**Description:** Renders a string at the specified coordinates.
Input: `s` (String) - The text to draw.
Input: `x`, `y` (Double) - Baseline coordinates.
Return: `Empty`.
**Observations:** Requires a font to be loaded first via `LoadFontFace`.
**Syntax:**
```vbscript
Dim img
Set img = Server.CreateObject("G3Image")
img.LoadFontFace "/fonts/roboto.ttf", 18
img.DrawString "Dynamic Image Text", 10, 30
```

## ServerObject: G3Image Function: RenderViaTemp(format, quality)
**Keywords:** render, temporary, output, bytearray, buffer, binary
**Description:** Renders the image to a temporary file and returns its binary content.
Input: `format` (String) - "png" or "jpg".
Input: `quality` (Integer) - Optional. 1-100 for JPG.
Return: `Array` (VTArray) - Byte array of the image data.
**Observations:** Highly useful for serving images directly from memory/ASP code without persistent files.
**Syntax:**
```vbscript
Dim img, bytes
Set img = Server.CreateObject("G3Image")
img.New 100, 100
img.SetHexColor "#0000FF"
img.DrawRectangle 0,0,100,100
img.Fill
bytes = img.RenderViaTemp("png", 0)
Response.BinaryWrite bytes
```

## ServerObject: G3JSON Function: Parse(jsonStr)
**Keywords:** json, parse, decode, object, array, dictionary, deserialize
**Description:** Parses a JSON string into native VBScript structures (Dictionaries and Arrays).
Input: `jsonStr` (String) - The JSON encoded string.
Return: `Variant` - Returns a `Dictionary` for JSON objects, an `Array` for JSON arrays, or primitive types.
**Observations:** High-performance parser that bridges Go's JSON engine with the AxonVM object model.
**Syntax:**
```vbscript
Dim json, data
Set json = Server.CreateObject("G3JSON")
Set data = json.Parse("{""id"": 1, ""name"": ""Axon""}")
Response.Write data("name")
```

## ServerObject: G3JSON Function: Stringify(value)
**Keywords:** json, encode, stringify, serialize, object, array
**Description:** Converts a VBScript variable (including Dictionaries and Arrays) into a JSON string.
Input: `value` (Variant) - The data structure to serialize.
Return: `String` - Valid JSON string.
**Observations:** Recursively traverses Dictionaries and Arrays to build the JSON tree.
**Syntax:**
```vbscript
Dim json, data
Set json = Server.CreateObject("G3JSON")
Set data = Server.CreateObject("Scripting.Dictionary")
data.Add "status", "success"
data.Add "code", 200
Response.Write json.Stringify(data)
```

## ServerObject: G3JSON Function: LoadFile(path)
**Keywords:** json, file, load, read, parse, config
**Description:** Reads a JSON file from disk and parses it immediately.
Input: `path` (String) - Path to the .json file.
Return: `Variant` - The parsed data structure.
**Observations:** Resolves path relative to the web root. More efficient than manual Read + Parse.
**Syntax:**
```vbscript
Dim json, config
Set json = Server.CreateObject("G3JSON")
Set config = json.LoadFile("/config/app.json")
```

## ServerObject: G3Mail Function: AddAddress(addr)
**Keywords:** mail, email, recipient, add, to
**Description:** Adds an email address to the list of recipients.
Input: `addr` (String) - The recipient's email address.
Return: `Boolean` - True on success.
**Observations:** Can be called multiple times to add multiple recipients.
**Syntax:**
```vbscript
Dim mail
Set mail = Server.CreateObject("G3Mail")
mail.AddAddress "user@example.com"
```

## ServerObject: G3Mail Function: Send(to, subject, body)
**Keywords:** mail, email, send, smtp, network, transmission
**Description:** Sends the configured email via SMTP.
Input: `to` (String) - Optional. Primary recipient.
Input: `subject` (String) - Optional. Email subject.
Input: `body` (String) - Optional. Email message body.
Return: `Variant` - True on success, or an Error String on failure.
**Observations:** If SMTP credentials (host, port, user, pass) are not set on the object, it automatically looks for `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS` environment variables.
**Syntax:**
```vbscript
Dim mail, result
Set mail = Server.CreateObject("G3Mail")
mail.From = "noreply@mydomain.com"
mail.Subject = "Alert"
mail.HTMLBody = "<h1>System Alert</h1>"
result = mail.Send()
If result = True Then
    Response.Write "Sent!"
Else
    Response.Write result ' Display error
End If
```

## ServerObject: G3Mail Function: Clear()
**Keywords:** mail, reset, clear, recipients, subject, body
**Description:** Resets all email fields (recipients, subject, body, etc.).
Return: `Boolean`.
**Observations:** Useful for reusing the same mail object for multiple distinct messages.
**Syntax:**
```vbscript
Dim mail
Set mail = Server.CreateObject("G3Mail")
' ... send first mail ...
mail.Clear()
' ... prepare and send second mail ...
```

## ServerObject: G3MD Function: Process(source)
**Keywords:** markdown, convert, html, gfm, processing, format
**Description:** Converts Markdown source text into sanitized HTML using GitHub Flavored Markdown (GFM).
Input: `source` (String) - The Markdown formatted string.
Return: `String` - The generated HTML.
**Observations:** High-performance conversion using the Goldmark engine. Supports tables, task lists, and auto-links. Options like `HardWraps` and `Unsafe` affect the output.
**Syntax:**
```vbscript
Dim md, html
Set md = Server.CreateObject("G3MD")
html = md.Process("# Hello Axon" & vbCrLf & "This is **Markdown**.")
Response.Write html
```

## ServerObject: G3PDF Function: New(orientation, unit, size)
**Keywords:** pdf, create, document, initialization, landscape, portrait, a4
**Description:** Initializes a new PDF document with specific orientation, measurement units, and page size.
Input: `orientation` (String) - "P" (Portrait) or "L" (Landscape). Default "P".
Input: `unit` (String) - "pt", "mm", "cm", "in". Default "mm".
Input: `size` (String) - "A3", "A4", "A5", "Letter", "Legal". Default "A4".
Return: `Object` - Returns the PDF object instance.
**Observations:** This is typically the first call after creating the object. It resets any existing document state.
**Syntax:**
```vbscript
Dim pdf
Set pdf = Server.CreateObject("G3PDF")
pdf.New "P", "mm", "A4"
```

## ServerObject: G3PDF Function: AddPage(orientation, size)
**Keywords:** pdf, page, add, layout, landscape, portrait
**Description:** Adds a new page to the PDF document.
Input: `orientation` (String) - Optional. "P" or "L".
Input: `size` (String) - Optional. "A4", "Letter", etc.
Return: `Boolean`.
**Observations:** If orientation or size are omitted, the document defaults are used.
**Syntax:**
```vbscript
Dim pdf
Set pdf = Server.CreateObject("G3PDF")
pdf.New "P", "mm", "A4"
pdf.AddPage "L", "A4" ' Add a landscape page
```

## ServerObject: G3PDF Function: SetFont(family, style, size)
**Keywords:** pdf, font, style, text, typography, size, bold, italic
**Description:** Sets the font used for subsequent text rendering.
Input: `family` (String) - "courier", "helvetica", "times", "symbol", "zapfdingbats".
Input: `style` (String) - "" (regular), "B" (bold), "I" (italic), "U" (underline).
Input: `size` (Double) - Font size in points.
Return: `Boolean`.
**Observations:** Standard PDF fonts are built-in. Custom TTF fonts can be loaded if supported by the underlying engine.
**Syntax:**
```vbscript
Dim pdf
Set pdf = Server.CreateObject("G3PDF")
pdf.New "P", "mm", "A4"
pdf.AddPage "", ""
pdf.SetFont "helvetica", "B", 16
```

## ServerObject: G3PDF Function: Cell(w, h, txt, border, ln, align, fill, link)
**Keywords:** pdf, cell, table, text, border, layout, alignment
**Description:** Prints a rectangular cell with text and borders.
Input: `w` (Double) - Cell width. If 0, extends to right margin.
Input: `h` (Double) - Cell height.
Input: `txt` (String) - Text to display.
Input: `border` (Variant) - 0 (no border), 1 (full border), or string like "L", "T", "R", "B".
Input: `ln` (Integer) - 0 (right), 1 (next line), 2 (below).
Input: `align` (String) - "L", "C", "R".
Input: `fill` (Boolean) - Whether to fill the background.
Input: `link` (Variant) - Internal link ID or external URL.
Return: `Boolean`.
**Observations:** The most common method for layout and structured text.
**Syntax:**
```vbscript
Dim pdf
Set pdf = Server.CreateObject("G3PDF")
pdf.New "P", "mm", "A4"
pdf.AddPage "", ""
pdf.SetFont "helvetica", "", 12
pdf.Cell 40, 10, "Hello World!", 1, 1, "C", False, ""
```

## ServerObject: G3PDF Function: WriteHTML(html)
**Keywords:** pdf, html, render, convert, formatting, tags
**Description:** Renders a limited subset of HTML tags directly into the PDF.
Input: `html` (String) - HTML string.
Return: `Empty`.
**Observations:** Supports tags like `<b>`, `<i>`, `<u>`, `<a>`, `<p>`, `<br>`, `<h1>`-`<h6>`, `<ul>`, `<li>`, `<table>`, `<tr>`, `<td>`. Useful for dynamic reporting.
**Syntax:**
```vbscript
Dim pdf
Set pdf = Server.CreateObject("G3PDF")
pdf.New "P", "mm", "A4"
pdf.AddPage "", ""
pdf.WriteHTML "<h1>Report</h1><p>This is <b>bold</b> text.</p>"
```

## ServerObject: G3PDF Function: Output(dest, fileName)
**Keywords:** pdf, output, save, download, stream, binary
**Description:** Generates the PDF and outputs it to the specified destination.
Input: `dest` (String) - "I" (inline response), "D" (download), "F" (save to file), "S" (return as string).
Input: `fileName` (String) - The name for the file or download.
Return: `Variant` - Depends on `dest`.
**Observations:** In AxonASP, "I" and "D" interact directly with the HTTP response buffer.
**Syntax:**
```vbscript
Dim pdf
Set pdf = Server.CreateObject("G3PDF")
pdf.New "P", "mm", "A4"
pdf.AddPage "", ""
pdf.SetFont "helvetica", "B", 16
pdf.Cell 0, 10, "PDF Output Test", 0, 1, "C", False, ""
pdf.Output "D", "test.pdf" ' Trigger download
```

## ServerObject: G3Template Function: Render(path, data)
**Keywords:** template, render, go-template, html, logic, binding
**Description:** Renders a Go-style HTML template file with provided data.
Input: `path` (String) - Path to the template file.
Input: `data` (Variant) - A Dictionary or Array containing data for the template.
Return: `String` - The rendered HTML/text content.
**Observations:** Leverages Go's `html/template` engine, providing secure, XSS-safe rendering.
**Syntax:**
```vbscript
Dim tmpl, data, html
Set tmpl = Server.CreateObject("G3Template")
Set data = Server.CreateObject("Scripting.Dictionary")
data.Add "Name", "Axon User"
html = tmpl.Render("/templates/welcome.html", data)
Response.Write html
```

## ServerObject: G3Zip Function: Create(path)
**Keywords:** zip, create, archive, compression, file
**Description:** Creates a new empty ZIP archive at the specified path.
Input: `path` (String) - The target .zip file path.
Return: `Boolean` - True on success.
**Observations:** If the file exists, it will be overwritten. Use `Open` to read existing archives.
**Syntax:**
```vbscript
Dim zip
Set zip = Server.CreateObject("G3Zip")
If zip.Create("/backups/data.zip") Then
    zip.AddFile "/data/report.txt", "report.txt"
    zip.Close
End If
```

## ServerObject: G3Zip Function: AddFile(source, nameInZip)
**Keywords:** zip, add, file, archive, compression
**Description:** Adds a file from the filesystem to the current ZIP archive.
Input: `source` (String) - The path to the physical file.
Input: `nameInZip` (String) - Optional. The path/name inside the ZIP archive.
Return: `Boolean`.
**Observations:** Requires the archive to be opened in write mode (via `Create`).
**Syntax:**
```vbscript
Dim zip
Set zip = Server.CreateObject("G3Zip")
zip.Create "archive.zip"
zip.AddFile "C:\docs\manual.pdf", "manuals/manual.pdf"
zip.Close
```

## ServerObject: G3Zip Function: ExtractAll(targetDir)
**Keywords:** zip, extract, unzip, decompress, archive
**Description:** Extracts all files from the current ZIP archive to a target directory.
Input: `targetDir` (String) - The destination directory.
Return: `Boolean`.
**Observations:** Requires the archive to be opened in read mode (via `Open`).
**Syntax:**
```vbscript
Dim zip
Set zip = Server.CreateObject("G3Zip")
If zip.Open("backup.zip") Then
    zip.ExtractAll "/restored_data"
    zip.Close
End If
```

## ServerObject: G3Zip Function: List()
**Keywords:** zip, list, files, content, archive, inventory
**Description:** Returns an array of filenames contained in the ZIP archive.
Return: `Array` (VTArray) - Zero-based array of strings.
**Observations:** Requires the archive to be opened in read mode.
**Syntax:**
```vbscript
Dim zip, files, i
Set zip = Server.CreateObject("G3Zip")
If zip.Open("data.zip") Then
    files = zip.List()
    For i = 0 To UBound(files)
        Response.Write "File: " & files(i) & "<br>"
    Next
    zip.Close
End If
```

## ServerObject: G3FC Function: Create(outputRel, sourcePaths, [password], [configDict])
**Keywords:** g3fc, archive, compress, encrypt, pack, bundle, fec, create
**Description:** Creates a new G3FC archive file from the specified source paths with optional AES-GCM encryption and forward error correction (FEC).
Input: `outputRel` (String) - The relative path to the output `.g3fc` archive file.
Input: `sourcePaths` (String or Array) - A single relative path or an array of paths (files or directories) to include in the archive.
Input: `password` (String) - Optional. The password used for AES-GCM encryption.
Input: `configDict` (Dictionary) - Optional. A dictionary specifying advanced settings like `CompressionLevel` (1-22), `GlobalCompression` (Boolean), `FECLevel` (Integer 1-100), and `SplitSize` (String, e.g., "100MB").
Return: `Boolean` - True if the archive was successfully created, False otherwise.
**Observations:** Highly efficient archive format. Paths are resolved relative to the web root.
**Syntax:**
```vbscript
Dim ax, config
Set ax = Server.CreateObject("G3FC")
Set config = Server.CreateObject("Scripting.Dictionary")
config.Add "CompressionLevel", 9
config.Add "FECLevel", 10
If ax.Create("/backups/data.g3fc", Array("/logs", "/images"), "mySecurePassword", config) Then
    Response.Write "Archive created successfully."
End If
```

## ServerObject: G3FC Function: Extract(archiveRel, outputRel, [password])
**Keywords:** g3fc, extract, uncompress, unpack, decrypt, restore
**Description:** Extracts all files from a G3FC archive into a specified destination directory.
Input: `archiveRel` (String) - The relative path to the `.g3fc` archive file.
Input: `outputRel` (String) - The relative path to the destination directory.
Input: `password` (String) - Optional. The password required if the archive is encrypted.
Return: `Boolean` - True on success, False on failure.
**Observations:** Automatically reassembles split archives and applies FEC recovery if data corruption is detected.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3FC")
If ax.Extract("/backups/data.g3fc", "/restore", "mySecurePassword") Then
    Response.Write "Extraction complete."
End If
```

## ServerObject: G3FC Function: List(archiveRel, [password], [unit], [details])
**Keywords:** g3fc, list, archive, content, files, inventory, inspect
**Description:** Returns an array of dictionaries representing the files contained within a G3FC archive.
Input: `archiveRel` (String) - The relative path to the `.g3fc` archive file.
Input: `password` (String) - Optional. The password required to read the index if it's encrypted.
Input: `unit` (String) - Optional. Size formatting unit ("KB", "MB", "GB", "TB"). Default is "KB".
Input: `details` (Boolean) - Optional. If True, includes extra metadata (Permissions, CreationTime, Checksum) in the output. Default is False.
Return: `Array` (VTArray) - An array of Dictionaries, each containing file metadata.
**Observations:** Reads only the archive index, making it very fast even for large archives.
**Syntax:**
```vbscript
Dim ax, files, i
Set ax = Server.CreateObject("G3FC")
files = ax.List("/backups/data.g3fc", "mySecurePassword")
For i = 0 To UBound(files)
    Response.Write files(i)("Path") & " - " & files(i)("FormattedSize") & "<br>"
Next
```

## ServerObject: G3FC Function: Info(archiveRel, outputRel, [password])
**Keywords:** g3fc, info, metadata, export, json, inspect, structure
**Description:** Exports the detailed file index metadata of a G3FC archive to a JSON file.
Input: `archiveRel` (String) - The relative path to the `.g3fc` archive file.
Input: `outputRel` (String) - The destination path for the generated `.json` file.
Input: `password` (String) - Optional. The password required if the index is encrypted.
Return: `Boolean` - True if the JSON was successfully exported, False otherwise.
**Observations:** Very useful for programmatic integration and analyzing the internal structure of chunked or split archives.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3FC")
If ax.Info("/backups/data.g3fc", "/logs/archive_info.json", "mySecurePassword") Then
    Response.Write "Metadata exported to JSON."
End If
```

## ServerObject: G3FC Function: Find(archiveRel, pattern, [password], [useRegex])
**Keywords:** g3fc, find, search, pattern, regex, archive, query
**Description:** Searches for specific files inside a G3FC archive based on a path pattern or regular expression.
Input: `archiveRel` (String) - The relative path to the `.g3fc` archive file.
Input: `pattern` (String) - The search string or regular expression pattern.
Input: `password` (String) - Optional. The password required if the index is encrypted.
Input: `useRegex` (Boolean) - Optional. Whether to treat the pattern as a regular expression. Default is False.
Return: `Array` (VTArray) - An array of Dictionaries containing `Path` and `Size` for each match.
**Observations:** Performs case-insensitive matching by default.
**Syntax:**
```vbscript
Dim ax, matches, i
Set ax = Server.CreateObject("G3FC")
matches = ax.Find("/backups/data.g3fc", "\.png$", "mySecurePassword", True)
For i = 0 To UBound(matches)
    Response.Write "Found image: " & matches(i)("Path") & "<br>"
Next
```

## ServerObject: G3FC Function: ExtractSingle(archiveRel, filePath, outputRel, [password])
**Keywords:** g3fc, extract, single, file, restore, specific
**Description:** Extracts a specific file or directory from a G3FC archive to a destination path.
Input: `archiveRel` (String) - The relative path to the `.g3fc` archive file.
Input: `filePath` (String) - The exact internal path of the file within the archive.
Input: `outputRel` (String) - The relative path to the destination directory.
Input: `password` (String) - Optional. The password required if the archive is encrypted.
Return: `Boolean` - True if the file was successfully extracted, False otherwise.
**Observations:** Faster than extracting the entire archive when only one file is needed.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3FC")
If ax.ExtractSingle("/backups/data.g3fc", "logs/app.log", "/restore", "mySecurePassword") Then
    Response.Write "File restored successfully."
End If
```

## ServerObject: G3TAR Function: Create(relPath)
**Keywords:** tar, create, archive, write, compress, tape
**Description:** Initializes a new TAR archive for writing.
Input: `relPath` (String) - The relative path to the new `.tar` file.
Return: `Boolean` - True on success, False if the path is invalid or cannot be created.
**Observations:** This must be called before using `AddFile`, `AddFolder`, or `AddText`. Replaces any existing file at the location.
**Syntax:**
```vbscript
Dim tar
Set tar = Server.CreateObject("G3TAR")
If tar.Create("/archives/backup.tar") Then
    tar.AddFile "/docs/readme.txt", "readme.txt"
    tar.Close()
End If
```

## ServerObject: G3TAR Function: Open(relPath)
**Keywords:** tar, open, read, load, archive, inspect
**Description:** Opens an existing TAR archive for reading and indexes its headers.
Input: `relPath` (String) - The relative path to the existing `.tar` file.
Return: `Boolean` - True on success, False if the file cannot be read.
**Observations:** Must be called before `List`, `ExtractAll`, `ExtractFile`, or `GetFileInfo`.
**Syntax:**
```vbscript
Dim tar
Set tar = Server.CreateObject("G3TAR")
If tar.Open("/archives/backup.tar") Then
    Response.Write "Archive has " & tar.Count & " entries."
    tar.Close()
End If
```

## ServerObject: G3TAR Function: AddFile(sourceRelPath, [nameInTar])
**Keywords:** tar, add, file, append, include
**Description:** Streams a single file into the currently open TAR archive.
Input: `sourceRelPath` (String) - The path to the physical file to add.
Input: `nameInTar` (String) - Optional. The internal path/name to use inside the archive. If omitted, the file's base name is used.
Return: `Boolean` - True on success.
**Observations:** Requires the archive to be opened in write mode via `Create`.
**Syntax:**
```vbscript
Dim tar
Set tar = Server.CreateObject("G3TAR")
tar.Create "/archives/data.tar"
tar.AddFile "/images/photo.jpg", "assets/photo.jpg"
tar.Close()
```

## ServerObject: G3TAR Function: AddFolder(sourceRelPath, [nameInTar])
**Keywords:** tar, add, folder, directory, recursive, batch
**Description:** Recursively adds an entire folder and its contents into the current TAR archive.
Input: `sourceRelPath` (String) - The path to the physical directory to add.
Input: `nameInTar` (String) - Optional. The base internal path inside the archive.
Return: `Boolean` - True on success.
**Observations:** Maintains the internal folder hierarchy relative to the source directory.
**Syntax:**
```vbscript
Dim tar
Set tar = Server.CreateObject("G3TAR")
tar.Create "/archives/project.tar"
tar.AddFolder "/my_app/src", "src"
tar.Close()
```

## ServerObject: G3TAR Function: AddFiles(paths, [prefix])
**Keywords:** tar, add, files, batch, array, list
**Description:** Adds multiple files from an array to the current TAR archive with an optional common prefix.
Input: `paths` (Array) - An array of relative file paths.
Input: `prefix` (String) - Optional. The internal directory prefix for all added files.
Return: `Boolean` - True on success.
**Observations:** Efficient for packaging specific lists of files.
**Syntax:**
```vbscript
Dim tar
Set tar = Server.CreateObject("G3TAR")
tar.Create "/archives/assets.tar"
tar.AddFiles Array("/img/1.png", "/img/2.png"), "images"
tar.Close()
```

## ServerObject: G3TAR Function: AddText(nameInTar, content)
**Keywords:** tar, add, text, string, memory, dynamic
**Description:** Writes dynamically generated text directly into the TAR archive as a file.
Input: `nameInTar` (String) - The internal path and filename in the archive.
Input: `content` (String) - The raw text content to write.
Return: `Boolean` - True on success.
**Observations:** Avoids the need to create temporary files on disk before archiving them.
**Syntax:**
```vbscript
Dim tar
Set tar = Server.CreateObject("G3TAR")
tar.Create "/archives/logs.tar"
tar.AddText "summary.txt", "Archive generated successfully."
tar.Close()
```

## ServerObject: G3TAR Function: List()
**Keywords:** tar, list, index, entries, contents
**Description:** Retrieves a list of all file and directory names contained within the open TAR archive.
Return: `Array` (VTArray) - An array of strings representing the entry names.
**Observations:** Requires the archive to be opened via `Open`.
**Syntax:**
```vbscript
Dim tar, items, i
Set tar = Server.CreateObject("G3TAR")
If tar.Open("/archives/data.tar") Then
    items = tar.List()
    For i = 0 To UBound(items)
        Response.Write items(i) & "<br>"
    Next
    tar.Close()
End If
```

## ServerObject: G3TAR Function: ExtractAll(targetRelPath)
**Keywords:** tar, extract, unzip, untar, unpack, restore
**Description:** Extracts all contents of the TAR archive into a target directory.
Input: `targetRelPath` (String) - The destination directory.
Return: `Boolean` - True on success.
**Observations:** Automatically creates the destination directories if they don't exist. Has built-in protection against path traversal (Zip Slip) attacks.
**Syntax:**
```vbscript
Dim tar
Set tar = Server.CreateObject("G3TAR")
If tar.Open("/archives/data.tar") Then
    tar.ExtractAll "/restored_files"
    tar.Close()
End If
```

## ServerObject: G3TAR Function: ExtractFile(entryName, targetRelPath)
**Keywords:** tar, extract, single, file, restore
**Description:** Extracts a single specific file from the open TAR archive.
Input: `entryName` (String) - The exact internal path/name of the file in the archive.
Input: `targetRelPath` (String) - The destination directory where the file will be extracted.
Return: `Boolean` - True on success, False if the file was not found.
**Observations:** Efficient for plucking specific items from a large archive.
**Syntax:**
```vbscript
Dim tar
Set tar = Server.CreateObject("G3TAR")
If tar.Open("/archives/data.tar") Then
    If tar.ExtractFile("assets/photo.jpg", "/images") Then
        Response.Write "File extracted!"
    End If
    tar.Close()
End If
```

## ServerObject: G3TAR Function: GetFileInfo(name)
**Keywords:** tar, info, metadata, file, size, permissions
**Description:** Retrieves metadata for a specific entry inside the open TAR archive.
Input: `name` (String) - The internal entry name.
Return: `Dictionary` (VTNativeObject) - A dictionary with properties like `Name`, `Size`, `Mode`, `Modified`, and `IsDir`.
**Observations:** Useful for inspecting file details without extracting them.
**Syntax:**
```vbscript
Dim tar, info
Set tar = Server.CreateObject("G3TAR")
If tar.Open("/archives/data.tar") Then
    Set info = tar.GetFileInfo("readme.txt")
    If Not IsEmpty(info) Then
        Response.Write "Size: " & info("Size") & " bytes"
    End If
    tar.Close()
End If
```

## ServerObject: G3TAR Function: Close()
**Keywords:** tar, close, dispose, free, release
**Description:** Closes the open TAR archive and releases all file handles and resources.
Return: `Boolean` - Always True.
**Observations:** MUST be called after finishing operations to prevent file locks.
**Syntax:**
```vbscript
Dim tar
Set tar = Server.CreateObject("G3TAR")
tar.Create "test.tar"
tar.Close()
```

## ServerObject: G3ZLIB Function: Compress(input, [level])
**Keywords:** zlib, compress, deflate, encode, bytes, string
**Description:** Compresses a string or byte array using the zlib algorithm.
Input: `input` (String or Array) - The plain text string or byte array (0-255) to compress.
Input: `level` (Integer) - Optional. The compression level (e.g., 1 for speed, 9 for max compression). Default is -1 (DefaultCompression).
Return: `Array` (VTArray) - A byte array containing the compressed data.
**Observations:** Standard zlib format.
**Syntax:**
```vbscript
Dim z, compressed
Set z = Server.CreateObject("G3ZLIB")
compressed = z.Compress("This is a long text to be compressed...", 9)
```

## ServerObject: G3ZLIB Function: Decompress(input)
**Keywords:** zlib, decompress, inflate, decode, bytes
**Description:** Decompresses a zlib-compressed byte array back into its original raw bytes.
Input: `input` (String or Array) - The compressed data.
Return: `Array` (VTArray) - A byte array containing the raw decompressed bytes.
**Observations:** Returns an empty result if the input is corrupted or not a valid zlib stream.
**Syntax:**
```vbscript
Dim z, rawBytes
Set z = Server.CreateObject("G3ZLIB")
rawBytes = z.Decompress(compressedDataArray)
```

## ServerObject: G3ZLIB Function: DecompressText(input)
**Keywords:** zlib, decompress, text, string, decode
**Description:** Decompresses a zlib-compressed byte array directly into a UTF-8 string.
Input: `input` (String or Array) - The compressed data.
Return: `String` - The decompressed plain text.
**Observations:** Alias `DecompressString` is also available. Assumes the original data was text.
**Syntax:**
```vbscript
Dim z, plainText
Set z = Server.CreateObject("G3ZLIB")
plainText = z.DecompressText(compressedDataArray)
Response.Write plainText
```

## ServerObject: G3ZLIB Function: CompressMany(input, [level])
**Keywords:** zlib, compress, array, batch, multiple
**Description:** Compresses an array of inputs in a single operation.
Input: `input` (Array) - An array containing strings or byte arrays.
Input: `level` (Integer) - Optional. The compression level.
Return: `Array` (VTArray) - An array containing the corresponding compressed byte arrays.
**Observations:** Extremely fast for batch operations.
**Syntax:**
```vbscript
Dim z, results
Set z = Server.CreateObject("G3ZLIB")
results = z.CompressMany(Array("Text 1", "Text 2", "Text 3"))
```

## ServerObject: G3ZLIB Function: DecompressMany(input)
**Keywords:** zlib, decompress, array, batch, multiple
**Description:** Decompresses an array of zlib-compressed byte arrays.
Input: `input` (Array) - An array of compressed byte arrays.
Return: `Array` (VTArray) - An array of raw decompressed byte arrays.
**Observations:** Fails and returns empty if any item in the array is invalid.
**Syntax:**
```vbscript
Dim z, originals
Set z = Server.CreateObject("G3ZLIB")
originals = z.DecompressMany(compressedArrays)
```

## ServerObject: G3ZLIB Function: CompressFile(sourcePath, targetPath, [level])
**Keywords:** zlib, compress, file, stream, encode
**Description:** Compresses an entire file directly from the filesystem to a target path using zlib.
Input: `sourcePath` (String) - The file to read from.
Input: `targetPath` (String) - The file to write the compressed output to.
Input: `level` (Integer) - Optional. The compression level.
Return: `Boolean` - True on success.
**Observations:** Streams data internally to maintain very low memory usage, even for gigabyte-sized files.
**Syntax:**
```vbscript
Dim z
Set z = Server.CreateObject("G3ZLIB")
If z.CompressFile("/logs/app.log", "/logs/app.log.zz", 9) Then
    Response.Write "File compressed."
End If
```

## ServerObject: G3ZLIB Function: DecompressFile(sourcePath, targetPath)
**Keywords:** zlib, decompress, file, stream, decode
**Description:** Decompresses a zlib-compressed file from the filesystem directly to a target file.
Input: `sourcePath` (String) - The compressed file.
Input: `targetPath` (String) - The destination file.
Return: `Boolean` - True on success.
**Observations:** Highly memory efficient. Ensure the output path directory exists or the engine will attempt to create it.
**Syntax:**
```vbscript
Dim z
Set z = Server.CreateObject("G3ZLIB")
If z.DecompressFile("/logs/app.log.zz", "/logs/app_restored.log") Then
    Response.Write "File restored."
End If
```

## ServerObject: G3ZLIB Function: Clear()
**Keywords:** zlib, clear, reset, initialize, dispose
**Description:** Clears any internal error state from the zlib instance.
Return: `Boolean` - Always True.
**Observations:** Useful if you are reusing the object after catching an error via `LastError`.
**Syntax:**
```vbscript
Dim z
Set z = Server.CreateObject("G3ZLIB")
z.Clear()
```

## ServerObject: G3ZSTD Function: SetLevel(level)
**Keywords:** zstd, level, compression, speed, configuration
**Description:** Sets the default Zstandard compression level for subsequent operations.
Input: `level` (Integer) - The Zstd compression level (ranges from -5 for fastest to 22 for maximum compression). Default is 3.
Return: `Boolean` - True if the level is valid and was set.
**Observations:** Also available as `SetCompressionLevel`. Re-initializes the internal encoder.
**Syntax:**
```vbscript
Dim z
Set z = Server.CreateObject("G3ZSTD")
z.SetLevel 10
```

## ServerObject: G3ZSTD Function: Compress(input, [level])
**Keywords:** zstd, compress, encode, bytes, string, high-ratio
**Description:** Compresses a string or byte array using Zstandard.
Input: `input` (String or Array) - The plain text string or byte array to compress.
Input: `level` (Integer) - Optional. The compression level. Overrides the default level for this call.
Return: `Array` (VTArray) - A byte array containing the compressed data.
**Observations:** Zstandard offers extremely fast decompression regardless of the compression level.
**Syntax:**
```vbscript
Dim z, compressed
Set z = Server.CreateObject("G3ZSTD")
compressed = z.Compress("Data to compress with zstd")
```

## ServerObject: G3ZSTD Function: Decompress(input)
**Keywords:** zstd, decompress, decode, bytes, inflate
**Description:** Decompresses a Zstd-compressed byte array back into its original raw bytes.
Input: `input` (String or Array) - The compressed data.
Return: `Array` (VTArray) - A byte array containing the raw decompressed bytes.
**Observations:** Fast execution.
**Syntax:**
```vbscript
Dim z, rawBytes
Set z = Server.CreateObject("G3ZSTD")
rawBytes = z.Decompress(compressedDataArray)
```

## ServerObject: G3ZSTD Function: DecompressText(input)
**Keywords:** zstd, decompress, text, string, decode
**Description:** Decompresses a Zstd-compressed byte array directly into a UTF-8 string.
Input: `input` (String or Array) - The compressed data.
Return: `String` - The decompressed plain text.
**Observations:** Alias `DecompressString` is also available.
**Syntax:**
```vbscript
Dim z, plainText
Set z = Server.CreateObject("G3ZSTD")
plainText = z.DecompressText(compressedDataArray)
Response.Write plainText
```

## ServerObject: G3ZSTD Function: CompressMany(input, [level])
**Keywords:** zstd, compress, array, batch, multiple
**Description:** Compresses an array of inputs in a single operation.
Input: `input` (Array) - An array containing strings or byte arrays.
Input: `level` (Integer) - Optional. The compression level.
Return: `Array` (VTArray) - An array containing the corresponding compressed byte arrays.
**Observations:** Maintains high throughput for processing chunks.
**Syntax:**
```vbscript
Dim z, results
Set z = Server.CreateObject("G3ZSTD")
results = z.CompressMany(Array("Chunk A", "Chunk B"))
```

## ServerObject: G3ZSTD Function: DecompressMany(input)
**Keywords:** zstd, decompress, array, batch, multiple
**Description:** Decompresses an array of Zstd-compressed byte arrays.
Input: `input` (Array) - An array of compressed byte arrays.
Return: `Array` (VTArray) - An array of raw decompressed byte arrays.
**Observations:** Fails and returns empty if any item is corrupted.
**Syntax:**
```vbscript
Dim z, originals
Set z = Server.CreateObject("G3ZSTD")
originals = z.DecompressMany(compressedArrays)
```

## ServerObject: G3ZSTD Function: CompressFile(sourcePath, targetPath, [level])
**Keywords:** zstd, compress, file, stream, encode
**Description:** Compresses an entire file directly from the filesystem to a target path using Zstandard.
Input: `sourcePath` (String) - The file to read from.
Input: `targetPath` (String) - The file to write the compressed output to.
Input: `level` (Integer) - Optional. The compression level.
Return: `Boolean` - True on success.
**Observations:** Highly efficient streaming implementation suitable for huge files.
**Syntax:**
```vbscript
Dim z
Set z = Server.CreateObject("G3ZSTD")
If z.CompressFile("/db/backup.sql", "/db/backup.sql.zst", 19) Then
    Response.Write "File compressed."
End If
```

## ServerObject: G3ZSTD Function: DecompressFile(sourcePath, targetPath)
**Keywords:** zstd, decompress, file, stream, decode
**Description:** Decompresses a Zstd-compressed file from the filesystem directly to a target file.
Input: `sourcePath` (String) - The compressed `.zst` file.
Input: `targetPath` (String) - The destination file.
Return: `Boolean` - True on success.
**Observations:** Very fast streaming decompression.
**Syntax:**
```vbscript
Dim z
Set z = Server.CreateObject("G3ZSTD")
If z.DecompressFile("/db/backup.sql.zst", "/db/restored.sql") Then
    Response.Write "File restored."
End If
```

## ServerObject: G3ZSTD Function: Clear()
**Keywords:** zstd, clear, reset, initialize, dispose
**Description:** Clears internal state and releases the active encoder/decoder.
Return: `Boolean` - Always True.
**Observations:** Helps force garbage collection of large pre-allocated encoder memory if the object is kept alive for a long time.
**Syntax:**
```vbscript
Dim z
Set z = Server.CreateObject("G3ZSTD")
z.Clear()
```

## ServerObject: G3DB Function: Open(driver, dsn)
**Keywords:** db, database, sql, connect, connection, open, mysql, postgres, mssql, sqlite, oracle
**Description:** Establishes a database connection pool.
Input: `driver` (String) - The database type (e.g., "mysql", "postgres", "mssql", "sqlite", "oracle").
Input: `dsn` (String) - The connection string/DSN.
Return: `Boolean` - True if the connection and ping were successful.
**Observations:** This is a true connection pool, not an ADODB wrapper. A 5-second timeout applies to the initial ping.
**Syntax:**
```vbscript
Dim db
Set db = Server.CreateObject("G3DB")
If db.Open("mysql", "user:pass@tcp(localhost:3306)/mydb") Then
    Response.Write "Connected!"
End If
```

## ServerObject: G3DB Function: OpenFromEnv([driver])
**Keywords:** db, connect, environment, toml, config, setup
**Description:** Opens a database connection using configuration settings defined in `axonasp.toml` or environment variables.
Input: `driver` (String) - Optional. The target driver (default: "mysql").
Return: `Boolean` - True if successfully connected.
**Observations:** The cleanest way to connect without hardcoding credentials in VBScript.
**Syntax:**
```vbscript
Dim db
Set db = Server.CreateObject("G3DB")
If db.OpenFromEnv("postgres") Then
    Response.Write "Connected via config!"
End If
```

## ServerObject: G3DB Function: Query(sql, [params...])
**Keywords:** db, query, select, recordset, fetch, rows
**Description:** Executes a SELECT query with optional parameters and returns a forward-only `G3DBResultSet`.
Input: `sql` (String) - The SQL statement. Use `?` for placeholders regardless of the backend database.
Input: `params...` (Variadic) - The values to safely bind to the `?` placeholders.
Return: `Object` - A `G3DBResultSet` instance, or `Empty` on failure.
**Observations:** Completely protects against SQL injection. Automatically translates `?` placeholders into the backend's required format (e.g., `$1` for Postgres, `@p1` for MSSQL).
**Syntax:**
```vbscript
Dim db, rs
Set db = Server.CreateObject("G3DB")
db.OpenFromEnv()
Set rs = db.Query("SELECT id, name FROM users WHERE active = ? AND role = ?", 1, "admin")
Do While Not rs.EOF
    Response.Write rs("name") & "<br>"
    rs.MoveNext
Loop
rs.Close()
```

## ServerObject: G3DB Function: QueryRow(sql, [params...])
**Keywords:** db, query, select, single, row, fetch
**Description:** Executes a SELECT statement that is expected to return at most one row, yielding a `G3DBRow` object.
Input: `sql` (String) - The SQL statement.
Input: `params...` (Variadic) - Placeholder values.
Return: `Object` - A `G3DBRow` instance.
**Observations:** Highly optimized for single-record lookups (e.g., fetching a user by ID). The returned row object must be consumed via `Scan` or `ScanMap`.
**Syntax:**
```vbscript
Dim db, row, dict
Set db = Server.CreateObject("G3DB")
db.OpenFromEnv()
Set row = db.QueryRow("SELECT name, email FROM users WHERE id = ?", 42)
Set dict = row.ScanMap("Name", "Email")
If Not IsEmpty(dict) Then
    Response.Write "User: " & dict("Name")
End If
```

## ServerObject: G3DB Function: Exec(sql, [params...])
**Keywords:** db, execute, insert, update, delete, modification
**Description:** Executes an INSERT, UPDATE, or DELETE statement.
Input: `sql` (String) - The SQL statement.
Input: `params...` (Variadic) - Placeholder values.
Return: `Object` - A `G3DBResult` providing `LastInsertId` and `RowsAffected`.
**Observations:** Returns `Empty` on execution failure. Always use this for data modifications.
**Syntax:**
```vbscript
Dim db, result
Set db = Server.CreateObject("G3DB")
db.OpenFromEnv()
Set result = db.Exec("UPDATE users SET last_login = ? WHERE id = ?", Now(), 42)
If Not IsEmpty(result) Then
    Response.Write "Updated " & result.RowsAffected & " rows."
End If
```

## ServerObject: G3DB Function: Prepare(sql)
**Keywords:** db, prepare, statement, performance, batch
**Description:** Compiles a SQL query into a reusable prepared statement (`G3DBStatement`).
Input: `sql` (String) - The SQL template.
Return: `Object` - A `G3DBStatement` instance.
**Observations:** Highly recommended when executing the exact same query repeatedly with different parameters. The statement supports `Query`, `QueryRow`, and `Exec`.
**Syntax:**
```vbscript
Dim db, stmt, result, i
Set db = Server.CreateObject("G3DB")
db.OpenFromEnv()
Set stmt = db.Prepare("INSERT INTO logs (level, msg) VALUES (?, ?)")
For i = 1 To 100
    stmt.Exec "info", "Log entry " & i
Next
stmt.Close()
```

## ServerObject: G3DB Function: BeginTx([timeoutSeconds], [readOnly])
**Keywords:** db, transaction, begin, commit, rollback, isolation
**Description:** Starts a database transaction (`G3DBTransaction`).
Input: `timeoutSeconds` (Integer) - Optional. The transaction timeout in seconds.
Input: `readOnly` (Boolean) - Optional. Suggests a read-only transaction if supported by the driver.
Return: `Object` - A `G3DBTransaction` object.
**Observations:** You can also use `Begin()`, `BeginTrans()`, or `BeginTransaction()` with zero parameters. Uncommitted transactions are automatically rolled back when the script ends.
**Syntax:**
```vbscript
Dim db, tx
Set db = Server.CreateObject("G3DB")
db.OpenFromEnv()
Set tx = db.BeginTx(10, False)
tx.Exec "UPDATE accounts SET balance = balance - 100 WHERE id = 1"
tx.Exec "UPDATE accounts SET balance = balance + 100 WHERE id = 2"
If tx.Commit() Then
    Response.Write "Transfer successful."
Else
    tx.Rollback()
    Response.Write "Transfer failed."
End If
```

## ServerObject: G3DB Function: SetMaxOpenConns(n)
**Keywords:** db, connection, pool, limit, tuning, performance
**Description:** Sets the maximum number of open connections allowed to the database.
Input: `n` (Integer) - The maximum number of connections.
Return: `Empty`.
**Observations:** 0 means unlimited. Good for preventing database exhaustion.
**Syntax:**
```vbscript
Dim db
Set db = Server.CreateObject("G3DB")
db.SetMaxOpenConns 50
db.OpenFromEnv()
```

## ServerObject: G3DB Function: SetMaxIdleConns(n)
**Keywords:** db, connection, pool, idle, tuning
**Description:** Sets the maximum number of connections allowed to remain in the idle pool.
Input: `n` (Integer) - The idle connection limit.
Return: `Empty`.
**Observations:** Default usually depends on the driver.
**Syntax:**
```vbscript
Dim db
Set db = Server.CreateObject("G3DB")
db.SetMaxIdleConns 10
db.OpenFromEnv()
```

## ServerObject: G3DB Function: SetConnMaxLifetime(seconds)
**Keywords:** db, connection, pool, lifetime, timeout, refresh
**Description:** Sets the maximum amount of time a connection may be reused.
Input: `seconds` (Integer) - Max connection lifetime in seconds.
Return: `Empty`.
**Observations:** Expired connections are transparently closed and replaced.
**Syntax:**
```vbscript
Dim db
Set db = Server.CreateObject("G3DB")
db.SetConnMaxLifetime 3600 ' 1 hour
db.OpenFromEnv()
```

## ServerObject: G3DB Function: SetConnMaxIdleTime(seconds)
**Keywords:** db, connection, pool, idle, timeout, refresh
**Description:** Sets the maximum amount of time a connection may remain idle before being closed.
Input: `seconds` (Integer) - Idle timeout in seconds.
Return: `Empty`.
**Observations:** Helps free resources during low traffic periods.
**Syntax:**
```vbscript
Dim db
Set db = Server.CreateObject("G3DB")
db.SetConnMaxIdleTime 300 ' 5 minutes
db.OpenFromEnv()
```

## ServerObject: G3DB Function: Stats()
**Keywords:** db, pool, statistics, metrics, monitoring, connections
**Description:** Returns real-time statistics regarding the database connection pool.
Return: `Dictionary` (VTNativeObject) - Contains keys like `MaxOpenConnections`, `OpenConnections`, `InUse`, `Idle`, and `WaitCount`.
**Observations:** Extremely useful for building APM dashboards or monitoring pool exhaustion.
**Syntax:**
```vbscript
Dim db, st
Set db = Server.CreateObject("G3DB")
db.OpenFromEnv()
Set st = db.Stats()
Response.Write "Active connections: " & st("InUse")
```

## ServerObject: G3DB Function: GetLastError()
**Keywords:** db, error, message, diagnostic, debug, exception
**Description:** Retrieves the last error message emitted by a database operation.
Return: `String` - The error message.
**Observations:** Can also be called via the `LastError` property. Useful when a `Query` or `Exec` returns `Empty`.
**Syntax:**
```vbscript
Dim db, rs
Set db = Server.CreateObject("G3DB")
db.OpenFromEnv()
Set rs = db.Query("INVALID SQL SYNTAX")
If IsEmpty(rs) Then
    Response.Write "Error: " & db.GetLastError()
End If
```

## ServerObject: G3Axon.Functions Function: AxRuntimeInfo()
**Keywords:** runtime, diagnostics, info, system info, phpinfo, server info, configuration, memory, environment, go runtime, platform, version, report, debug, inspect
**Description:** Returns a multi-section plain-text diagnostic report modelled on PHP's `phpinfo()` output. The report is a single `String` value covering: engine name, runtime version, Go runtime version, platform (`GOOS/GOARCH`), CPU core count, GOMAXPROCS; server context (hostname, PID, current working directory, executable path, user home directory, resolved config file path, cache directory, path separator, integer size); a live memory snapshot (alloc bytes, total alloc, sys, heap objects, GC runs); all configuration keys and values currently loaded from `config/axonasp.toml`; optional script-timeout from the ASP server settings; and a full AxonASP legal attribution block (copyright, MPL 2.0 notice, attribution clause, contribution policy). Takes no arguments. Return: `String`.
**Observations:** This function calls `runtime.ReadMemStats` on every invocation — avoid tight-loop calls in hot code paths. When the config file cannot be located, the CONFIGURATION section reports `Config status: not loaded`. The legal attribution block at the end always includes the verbatim copyright, project URL, and license text required by the AxonASP license. On the HTTP server the `AXONASP SERVER SETTINGS` section also prints the configured script timeout. Output is safe to wrap in a `<pre>` tag after HTML-encoding it with `Server.HTMLEncode`. Method name is case-insensitive.
**Syntax:**
```vbscript
Dim ax, report
Set ax = Server.CreateObject("G3Axon.Functions")
report = ax.AxRuntimeInfo()
Response.Write "<pre>" & Server.HTMLEncode(report) & "</pre>"
Set ax = Nothing
```

## ServerObject: G3Axon.Functions Function: AxUserHomeDirPath()
**Keywords:** home directory, user home, home path, user directory, USERPROFILE, HOME, os user, filesystem, path, current user, environment variable, home folder
**Description:** Returns the home directory of the OS user that is running the AxonASP process as a `String`. The resolution order is: (1) `os.UserHomeDir()`, (2) the `HomeDir` field of `user.Current()`, (3) the `USERPROFILE` environment variable (Windows), (4) the `HOME` environment variable (Unix). Returns an empty string when none of these sources produce a usable path. Takes no arguments. Return: `String`.
**Observations:** On Windows the result is typically `C:\Users\<username>`. On Linux/macOS it is `/home/<username>` or `/root` for root accounts. When the process runs under a restricted service account with no assigned home directory, all fallbacks may fail and the function returns `""`. Always check for an empty string before constructing paths from this value. Method name is case-insensitive.
**Syntax:**
```vbscript
Dim ax, homePath
Set ax = Server.CreateObject("G3Axon.Functions")
homePath = ax.AxUserHomeDirPath()
If homePath <> "" Then
    Response.Write "User home: " & homePath
Else
    Response.Write "Home directory not available"
End If
Set ax = Nothing
```

## ServerObject: G3Axon.Functions Function: AxUserConfigDirPath()
**Keywords:** config path, configuration file, axonasp.toml, config directory, settings file, toml, viper, config location, resolved path, absolute path
**Description:** Returns the resolved absolute path to the AxonASP configuration file `config/axonasp.toml` as a `String`. The resolution probes the following candidates in order and returns the first one that exists on disk: (1) `config/axonasp.toml` relative to the current working directory, (2) `../config/axonasp.toml`, (3) `config/axonasp.toml` relative to the running executable's directory. If none of the candidates exist the function still returns an absolute path built from CWD + `config/axonasp.toml` as a best-effort fallback. Takes no arguments. Return: `String`.
**Observations:** The returned path always uses the OS-native path separator and is always absolute. This function does not create the file if it is missing — use `AxGetConfig` to read values from it. Useful for logging, displaying the active config location in admin panels, or verifying config placement during deployment checks. Method name is case-insensitive.
**Syntax:**
```vbscript
Dim ax, cfgPath
Set ax = Server.CreateObject("G3Axon.Functions")
cfgPath = ax.AxUserConfigDirPath()
Response.Write "Config file: " & cfgPath
Set ax = Nothing
```

## ServerObject: G3Axon.Functions Function: AxCacheDirPath()
**Keywords:** cache directory, cache path, temp directory, .temp/cache, script cache, bytecode cache, cache folder, absolute path, temp folder
**Description:** Returns the absolute path to the AxonASP bytecode/script cache directory (`.temp/cache/`) as a `String`. The path is always absolute and always ends with the OS path separator character (backslash on Windows, forward slash on POSIX). Takes no arguments. Return: `String`.
**Observations:** The directory `.temp/cache` is resolved relative to the current working directory at call time. The function does not verify whether the directory actually exists — it only resolves and returns the path. The trailing separator is always included, so you can concatenate a file name directly without adding a separator yourself. Method name is case-insensitive.
**Syntax:**
```vbscript
Dim ax, cachePath
Set ax = Server.CreateObject("G3Axon.Functions")
cachePath = ax.AxCacheDirPath()
Response.Write "Cache directory: " & cachePath
Set ax = Nothing
```

## ServerObject: G3Axon.Functions Function: AxIsPathSeparator(character)
**Keywords:** path separator, file separator, directory separator, slash, backslash, os separator, filesystem, path character, check separator, validate path char
**Description:** Returns `True` if the supplied single-character `String` is a valid path separator for the current operating system, or `False` otherwise. On Windows both `\` and `/` are recognized as separators. On POSIX (Linux, macOS) only `/` qualifies. Input: `character` (String, one character). Return: `Boolean`.
**Observations:** The check is delegated to Go's `os.IsPathSeparator`, which matches the OS-level definition exactly. If the argument is absent, empty, or longer than one character the function returns `False`. Multi-byte Unicode rune sequences longer than one rune also return `False`. Method name is case-insensitive.
**Syntax:**
```vbscript
Dim ax
Set ax = Server.CreateObject("G3Axon.Functions")
If ax.AxIsPathSeparator("/") Then
    Response.Write "/ is a path separator on this OS"
End If
If Not ax.AxIsPathSeparator("a") Then
    Response.Write "a is NOT a path separator"
End If
Set ax = Nothing
```

## ServerObject: G3Axon.Functions Function: AxChangeTimes(path, accessTime, modifiedTime)
**Keywords:** file times, change timestamp, chtimes, access time, modification time, mtime, atime, unix timestamp, file metadata, touch, update file date
**Description:** Sets the access time and modification time of the file or directory at `path` using Unix epoch integers. Input: `path` (String) — target file or directory path. `accessTime` (Integer/Long) — desired access timestamp in seconds since Unix epoch. `modifiedTime` (Integer/Long) — desired modification timestamp in seconds since Unix epoch. Return: `Boolean` — `True` on success, `False` on any error (file not found, permission denied, empty path, fewer than 3 arguments).
**Observations:** Internally calls Go's `os.Chtimes`. On Windows, changing `atime` may be ignored by some filesystems (e.g., NTFS with access-time tracking disabled). The function accepts standard VBScript integer or long values as timestamps; pass `-1` for the current time equivalent by feeding `AxTime()` result. Returns `False` if fewer than 3 arguments are provided. Method name is case-insensitive.
**Syntax:**
```vbscript
Dim ax, ok
Set ax = Server.CreateObject("G3Axon.Functions")
ok = ax.AxChangeTimes("C:\temp\sample.txt", 1700000000, 1700000001)
If ok Then
    Response.Write "Timestamps updated successfully"
Else
    Response.Write "Failed to update timestamps"
End If
Set ax = Nothing
```

## ServerObject: G3Axon.Functions Function: AxChangeMode(path, mode)
**Keywords:** chmod, file permissions, file mode, octal, permission bits, access rights, file security, unix permissions, filesystem, change permissions
**Description:** Changes the permission bits of the file at `path` using an octal text representation. Input: `path` (String) — target file path. `mode` (String) — octal permission text, for example `"0644"`, `"0755"`, or `"0600"`. Return: `Boolean` — `True` on success, `False` on any error (invalid path, invalid mode string, permission denied, fewer than 2 arguments).
**Observations:** The `mode` argument is parsed as a base-8 integer using Go's `strconv.ParseUint(modeText, 8, 32)`. Passing a decimal string like `"644"` (without a leading zero) is also accepted since the string is always parsed as octal. On Windows, `os.Chmod` only controls the read-only bit; full Unix permission semantics are not enforced by the OS. Returns `False` if the mode string is empty or cannot be parsed as a valid octal number. Method name is case-insensitive.
**Syntax:**
```vbscript
Dim ax, ok
Set ax = Server.CreateObject("G3Axon.Functions")
ok = ax.AxChangeMode("/var/www/data/file.txt", "0644")
If ok Then
    Response.Write "Permissions changed"
Else
    Response.Write "Could not change permissions"
End If
Set ax = Nothing
```

## ServerObject: G3Axon.Functions Function: AxCreateLink(sourcePath, linkPath)
**Keywords:** hard link, link, create link, os link, file link, hardlink, filesystem, duplicate entry, link file, create hard link
**Description:** Creates a hard link at `linkPath` pointing to the existing file at `sourcePath`. Input: `sourcePath` (String) — path to the existing source file. `linkPath` (String) — path where the hard link should be created. Return: `Boolean` — `True` on success, `False` on any error (source not found, destination already exists, cross-device link attempt, permission denied, empty path arguments, fewer than 2 arguments).
**Observations:** Calls Go's `os.Link`, which creates a hard link (not a symbolic link). Hard links share the same inode as the source; deleting one does not remove the other. Cross-filesystem or cross-volume hard links will always fail — both paths must reside on the same filesystem/volume. On Windows hard links are supported on NTFS but require appropriate privileges; some restricted server environments may deny creation. If `linkPath` already exists the function returns `False`. Method name is case-insensitive.
**Syntax:**
```vbscript
Dim ax, ok
Set ax = Server.CreateObject("G3Axon.Functions")
ok = ax.AxCreateLink("C:\temp\original.txt", "C:\temp\original.link")
If ok Then
    Response.Write "Hard link created"
Else
    Response.Write "Link creation failed (may be restricted on this OS or filesystem)"
End If
Set ax = Nothing
```

## ServerObject: G3Axon.Functions Function: AxChangeOwner(path, uid, gid)
**Keywords:** chown, change owner, file ownership, uid, gid, user id, group id, unix owner, file security, permission, owner change, os.chown
**Description:** Changes the owner user ID and group ID of the file at `path`. Input: `path` (String) — target file path. `uid` (Integer) — numeric user ID to assign as owner. `gid` (Integer) — numeric group ID to assign. Return: `Boolean` — `True` on success, `False` on any error (permission denied, unsupported platform, empty path, fewer than 3 arguments).
**Observations:** Delegates to Go's `os.Chown`. On **Windows**, `os.Chown` is a no-op and always returns an error, so this function always returns `False` on Windows regardless of the supplied IDs. On Linux/macOS the calling process must be root or have `CAP_CHOWN` capability to change ownership to an arbitrary UID/GID; unprivileged processes can only change the group if they own the file and are a member of the target group. Passing `0` for both `uid` and `gid` on a non-privileged context will also return `False`. Method name is case-insensitive.
**Syntax:**
```vbscript
Dim ax, ok
Set ax = Server.CreateObject("G3Axon.Functions")
ok = ax.AxChangeOwner("/var/www/data/file.txt", 1000, 1000)
If ok Then
    Response.Write "Owner changed successfully"
Else
    Response.Write "Could not change owner (expected on Windows or non-privileged environments)"
End If
Set ax = Nothing
```
