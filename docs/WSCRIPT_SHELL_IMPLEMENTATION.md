# WScript.Shell Implementation

## Overview

The WScript.Shell object provides access to the Windows Shell (cmd.exe on Windows, sh on Unix-like systems). It allows ASP applications to execute shell commands, both synchronously and asynchronously, with full support for input/output streaming.

## Full Compatibility

- **Windows**: 100% compatible with classic ASP WScript.Shell
- **Unix/Linux/macOS**: Fully functional with equivalent shell commands (sh instead of cmd.exe)
- **Cross-platform**: Automatic shell selection based on operating system

## Object Reference

### Creating a WScript.Shell Object

```vbscript
Set objShell =  Server.CreateObject("WScript.Shell")
```

Or use the shorthand:

```vbscript
Set objShell =  Server.CreateObject("Wscript.Shell")
Set objShell =  Server.CreateObject("Shell")
```

## Methods and Properties

### Run Method

Executes a shell command synchronously (with optional asynchronous execution).

**Syntax:**
```vbscript
intReturnCode = objShell.Run(strCommand, [intWindowStyle], [bWaitOnReturn])
```

**Parameters:**

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| strCommand | String | Command to execute | Required |
| intWindowStyle | Integer | Window display style (0-10) | 1 |
| bWaitOnReturn | Boolean | Wait for command completion | True |

**Window Styles:**
- 0 = Hidden
- 1 = Normal (default)
- 2 = Minimized
- 3 = Maximized
- 10 = Default (same as 1)

**Return Value:**
- Integer: Exit code of the executed command (0 = success)
- Note: Window style parameter is ignored on non-Windows systems

**Example:**

```vbscript
Set objShell = Server.CreateObject("WScript.Shell")

' Run command synchronously, wait for completion
intReturn = objShell.Run("ipconfig", 1, True)
Response.Write "Command returned: " & intReturn

' Run command asynchronously, don't wait
objShell.Run "notepad.exe", 1, False
```

### Exec Method

Executes a shell command asynchronously with full access to stdout, stderr, and stdin streams.

**Syntax:**
```vbscript
Set objExec = objShell.Exec(strCommand)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| strCommand | String | Command to execute (required) |

**Return Value:**
- WScriptExecObject: Object with access to input/output streams

**Example:**

```vbscript
Set objShell =  Server.CreateObject("WScript.Shell")
Set objExec = objShell.Exec("dir C:\")

' Read all output at once
strOutput = objExec.StdOut.ReadAll()
Response.Write "<pre>" & strOutput & "</pre>"
```

## WScriptExecObject

Returned by the **Exec()** method.

### Properties

#### Status

Returns the status of the process.

- **0** = Process is running
- **1** = Process has completed

```vbscript
Set objExec = objShell.Exec("ping google.com")
Response.Write "Status: " & objExec.Status
```

#### ExitCode

Returns the exit code of the completed process.

- **0** = Success
- **Non-zero** = Failure or specific error code
- **-1** = Process still running or error

```vbscript
Set objExec = objShell.Exec("dir")
objExec.WaitUntilDone()
Response.Write "Exit Code: " & objExec.ExitCode
```

#### ProcessID

Returns the process ID of the executing command.

```vbscript
Set objExec = objShell.Exec("notepad.exe")
Response.Write "Process ID: " & objExec.ProcessID
```

#### StdOut

TextStream object for reading standard output.

```vbscript
Set objExec = objShell.Exec("dir")
Set objStdOut = objExec.StdOut
While Not objStdOut.AtEndOfStream
    Response.Write objStdOut.ReadLine() & "<br>"
Wend
```

#### StdErr

TextStream object for reading standard error output.

```vbscript
Set objExec = objShell.Exec("command_with_error")
Set objStdErr = objExec.StdErr
If Not objStdErr.AtEndOfStream Then
    strError = objStdErr.ReadAll()
    Response.Write "Error: " & strError
End If
```

#### StdIn

TextStream object for writing to standard input.

```vbscript
Set objExec = objShell.Exec("sort")
objExec.StdIn.WriteLine("zebra")
objExec.StdIn.WriteLine("apple")
objExec.StdIn.WriteLine("banana")
objExec.StdIn.Close()
```

### Methods

#### WaitUntilDone

Waits for the process to complete.

**Syntax:**
```vbscript
objExec.WaitUntilDone([intTimeout])
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| intTimeout | Integer | Timeout in milliseconds (0 = wait indefinitely) |

**Return Value:**
- Boolean: True if process completed, False if timeout

**Example:**

```vbscript
Set objShell =  Server.CreateObject("WScript.Shell")
Set objExec = objShell.Exec("ping localhost")

' Wait up to 5 seconds for completion
bCompleted = objExec.WaitUntilDone(5000)

If bCompleted Then
    Response.Write "Exit Code: " & objExec.ExitCode
Else
    Response.Write "Command timed out"
    objExec.Terminate()
End If
```

#### Terminate

Terminates the running process.

**Syntax:**
```vbscript
objExec.Terminate()
```

**Example:**

```vbscript
Set objExec = objShell.Exec("long_running_command")
objExec.Terminate()
```

## TextStream Object

Used for reading/writing data to/from process streams.

### Properties

#### AtEndOfStream

Boolean indicating if the stream has reached the end.

```vbscript
Set objExec = objShell.Exec("dir")
Set objStdOut = objExec.StdOut

Do While Not objStdOut.AtEndOfStream
    Response.Write objStdOut.ReadLine() & "<br>"
Loop
```

### Methods

#### Read

Reads a specified number of characters from the stream.

**Syntax:**
```vbscript
strData = objStream.Read(numChars)
```

**Parameters:**
- numChars (Integer): Number of characters to read

**Return Value:**
- String: Up to numChars characters from the stream (or fewer if reaching end)

```vbscript
Set objExec = objShell.Exec("dir")
strFirstThousand = objExec.StdOut.Read(1000)
```

#### ReadLine

Reads one line (up to newline character) from the stream.

**Syntax:**
```vbscript
strLine = objStream.ReadLine()
```

**Return Value:**
- String: One line without the newline character

```vbscript
Do While Not objExec.StdOut.AtEndOfStream
    Response.Write objExec.StdOut.ReadLine() & "<br>"
Loop
```

#### ReadAll

Reads all remaining data from the stream.

**Syntax:**
```vbscript
strAllData = objStream.ReadAll()
```

**Return Value:**
- String: All remaining content from the stream

```vbscript
Set objExec = objShell.Exec("dir C:\")
strOutput = objExec.StdOut.ReadAll()
Response.Write "<pre>" & strOutput & "</pre>"
```

#### Write

Writes text to the stream (typically stdin).

**Syntax:**
```vbscript
objStream.Write(strText)
```

```vbscript
Set objExec = objShell.Exec("sort")
objExec.StdIn.Write("zebra")
objExec.StdIn.Write("apple")
```

#### WriteLine

Writes text with newline to the stream.

**Syntax:**
```vbscript
objStream.WriteLine(strText)
```

```vbscript
objExec.StdIn.WriteLine("line1")
objExec.StdIn.WriteLine("line2")
```

#### Close

Closes the stream.

**Syntax:**
```vbscript
objStream.Close()
```

```vbscript
objExec.StdIn.Close()  ' Terminate input to sort command
```

## GetEnv Method

Gets an environment variable value.

**Syntax:**
```vbscript
strValue = objShell.GetEnv(strVarName)
```

**Parameters:**
- strVarName (String): Name of environment variable

**Return Value:**
- String: Value of the environment variable (empty string if not found)

**Example:**

```vbscript
Set objShell = CreateObject("WScript.Shell")

' Get system environment variables
strPath = objShell.GetEnv("PATH")
strTemp = objShell.GetEnv("TEMP")
strOS = objShell.GetEnv("OS")

Response.Write "OS: " & strOS & "<br>"
Response.Write "TEMP: " & strTemp
```

## Practical Examples

### Example 1: Simple Command Execution

```vbscript
<%
Set objShell =  Server.CreateObject("WScript.Shell")
intCode = objShell.Run("notepad.exe", 1, False)  ' Open Notepad asynchronously
%>
```

### Example 2: Capture Command Output

```vbscript
<%
Set objShell =  Server.CreateObject("WScript.Shell")
Set objExec = objShell.Exec("dir C:\")

Response.Write "<h2>Directory Listing:</h2>"
Response.Write "<pre>" & objExec.StdOut.ReadAll() & "</pre>"
%>
```

### Example 3: Process Line-by-Line Output

```vbscript
<%
Set objShell =  Server.CreateObject("WScript.Shell")
Set objExec = objShell.Exec("ipconfig")

Response.Write "<h2>Network Configuration:</h2>"
Response.Write "<pre>"
Do While Not objExec.StdOut.AtEndOfStream
    Response.Write objExec.StdOut.ReadLine() & vbCrLf
Loop
Response.Write "</pre>"
%>
```

### Example 4: Send Input to Process

```vbscript
<%
Set objShell =  Server.CreateObject("WScript.Shell")
Set objExec = objShell.Exec("sort")

' Send unsorted data
objExec.StdIn.WriteLine("Zebra")
objExec.StdIn.WriteLine("Apple")
objExec.StdIn.WriteLine("Banana")
objExec.StdIn.Close()

' Read sorted output
Response.Write "<h2>Sorted Output:</h2>"
Response.Write "<pre>" & objExec.StdOut.ReadAll() & "</pre>"
%>
```

### Example 5: Error Handling

```vbscript
<%
Set objShell =  Server.CreateObject("WScript.Shell")
Set objExec = objShell.Exec("dir /nonexistent")

' Check for errors
If Not objExec.StdErr.AtEndOfStream Then
    Response.Write "Error: " & objExec.StdErr.ReadAll()
Else
    Response.Write "Output: " & objExec.StdOut.ReadAll()
End If
%>
```

### Example 6: Process with Timeout

```vbscript
<%
Set objShell =  Server.CreateObject("WScript.Shell")
Set objExec = objShell.Exec("ping localhost -n 10")

' Wait up to 10 seconds
If objExec.WaitUntilDone(10000) Then
    Response.Write "Exit Code: " & objExec.ExitCode & "<br>"
    Response.Write "Output: " & objExec.StdOut.ReadAll()
Else
    Response.Write "Command timed out, terminating..."
    objExec.Terminate()
End If
%>
```

## Platform-Specific Notes

### Windows
- Uses `cmd.exe /c` for command execution
- All shell commands work as expected
- Window styles are respected
- Full compatibility with Windows batch files

### Unix/Linux/macOS
- Uses `sh -c` for command execution
- Window styles are ignored (but don't cause errors)
- Shell commands should be POSIX-compatible
- Use appropriate Unix commands (ls instead of dir, etc.)

## Performance Considerations

1. **Do not create many concurrent processes** - each Exec() call starts a new process
2. **Close streams when done** - call `objExec.StdIn.Close()` after writing input
3. **ReadAll() vs ReadLine()** - use ReadLine() for large outputs to avoid memory issues
4. **Timeout handling** - always set reasonable timeouts for external commands

## Limitations

1. **Standard I/O Only** - Supports stdin, stdout, stderr only; not environment redirection
2. **No working directory** - Current working directory is inherited from server
3. **Security** - Be extremely careful with user input; always validate and sanitize
4. **Timeouts** - Some long-running commands may exceed server script timeout
5. **Output encoding** - Output is UTF-8 (may differ from Windows console default)

## Security Considerations

⚠️ **WARNING**: Executing shell commands based on user input is extremely dangerous!

**Always validate and sanitize user input:**

```vbscript
<%
' BAD: Don't do this!
userInput = Request.QueryString("cmd")
objShell.Run userInput  ' DANGEROUS!

' GOOD: Validate and use safe commands only
allowedCommands = Array("dir", "ping", "ipconfig")
userCommand = Request.QueryString("cmd")

Dim isAllowed
isAllowed = False
For Each cmd In allowedCommands
    If LCase(userCommand) = LCase(cmd) Then
        isAllowed = True
        Exit For
    End If
Next

If isAllowed Then
    objShell.Run userCommand
End If
%>
```

## See Also

- [AxExecute() Function](CUSTOM_FUNCTIONS.md#axexecute)
- [Scripting Objects Implementation](SCRIPTING_OBJECTS_IMPLEMENTATION.md)
