# Methods

## Overview

This page lists every method exposed by G3AXON.Functions in AxonASP.

## Method List
- Axarrayreverse: Returns a new array with elements in reverse order.
- Axbase64decode: Decodes a Base64 string back to plain text.
- Axbase64encode: Encodes a string into Base64.
- Axceil: Rounds a numeric value up to the nearest integer boundary.
- Axchangedir: Changes the current process working directory and returns success as Boolean.
- Axclearenvironment: Clears all process environment variables and returns True.
- Axcount: Returns the element count of a VB array value.
- Axctypealnum: Checks whether a string contains only alphanumeric ASCII characters.
- Axctypealpha: Checks whether a string contains only alphabetic ASCII letters.
- Axcurrentdir: Returns the current process working directory path.
- Axcurrentuser: Returns the current operating system user name.
- Axdate: Formats a date/time value using PHP-style format tokens, with optional timestamp argument.
- Axdirseparator: Returns the OS directory separator character.
- Axeffectiveuserid: Returns the effective user ID on Unix-like systems, or -1 on Windows.
- Axempty: Checks whether a value is considered empty (Empty, Null, empty string, zero, or False).
- Axenginename: Returns the AxonASP engine name string.
- Axenvironmentlist: Returns an array of environment entries in KEY=VALUE format.
- Axenvironmentvalue: Returns an environment variable value with optional default fallback when the variable is not found.
- Axexecutablepath: Returns the absolute path of the running AxonASP executable.
- Axexecute: Executes a system shell command and returns the combined stdout/stderr output as text.
- Axexplode: Splits a string into a VB array using a delimiter, with optional limit.
- Axfiltervalidateemail: Validates whether a string is a syntactically valid email address.
- Axfiltervalidateip: Validates whether a string is a valid IP address.
- Axfloatprecisiondigits: Returns the standard float precision digit count used by this library (15).
- Axfloor: Rounds a numeric value down to the nearest integer boundary.
- Axgenerateguid: Generates a random RFC 4122 version 4 GUID string.
- Axgetconfig: Reads one AxonASP configuration key from axonasp.toml (with env override support when enabled).
- Axgetconfigkeys: Returns an array with all available AxonASP configuration keys.
- Axgetdefaultcss: Reads and returns the configured default CSS file content.
- Axgetenv: Reads an environment variable by name and returns its value.
- Axgetlogo: Loads the configured logo file and returns it as a Base64 data URI.
- Axgetremotefile: Fetches a remote HTTP/HTTPS resource and returns the response body text on success.
- Axhash: Returns a hash digest for the selected algorithm (sha256, sha1, or md5).
- Axhextorgb: Converts an HTML hex color string into an rgb(r,g,b) string.
- Axhostnamevalue: Returns the current machine host name.
- Axhtmlspecialchars: Escapes special HTML characters in a string.
- Aximplode: Joins array items into a single string using a glue separator.
- Axintegermax: Returns the maximum native integer value for the current platform.
- Axintegermin: Returns the minimum native integer value for the current platform.
- Axintegersizebytes: Returns the native integer size in bytes for the current platform.
- Axisfloat: Checks whether the value type is Double (floating-point).
- Axisint: Checks whether the value type is Integer.
- Axisset: Checks whether a value is set (not Empty and not Null).
- Axlastmodified: Returns the Unix timestamp of the mapped script path last modification time.
- Axmax: Returns the largest numeric value from the provided arguments.
- Axmd5: Returns the MD5 hash of a string as lowercase hexadecimal.
- Axmin: Returns the smallest numeric value from the provided arguments.
- Axnl2br: Converts line breaks (CRLF/LF/CR) to HTML <br> tags.
- Axnumberformat: Formats a numeric value with configurable decimal places, decimal point, and thousands separator.
- Axpad: Pads a string to a target length (left, right, or both sides).
- Axpathlistseparator: Returns the OS path-list separator character used in PATH-like variables.
- Axpi: Returns the mathematical constant PI.
- Axplatformbits: Returns the native integer size in bits for the current platform.
- Axprocessid: Returns the current process ID (PID).
- Axrand: Returns a random integer (single bound or range, depending on arguments).
- Axrange: Builds and returns a numeric array range from start to end with optional step.
- Axrawurldecode: Decodes URL-encoded text after converting plus signs to spaces.
- Axrepeat: Repeats a string a specified number of times.
- Axrgbtohex: Converts RGB channel values to an HTML hexadecimal color string.
- Axsha1: Returns the SHA-1 hash of a string as lowercase hexadecimal.
- Axshutdownaxonaspserver: Shuts down the AxonASP server process when the shutdown function is enabled in configuration.
- Axsmallestfloatvalue: Returns the smallest non-zero positive float64 value supported by the runtime.
- Axstringgetcsv: Parses one CSV row string and returns the values as an array.
- Axstringreplace: Replaces all occurrences of a substring in a source string.
- Axstriptags: Removes HTML/XML-like tags from a string.
- Axsysteminfo: Returns system information (full string or selected mode such as OS, hostname, runtime version, or architecture).
- Axtime: Returns the current Unix timestamp in seconds.
- Axtrim: Trims whitespace or custom characters from both ends of a string.
- Axucfirst: Uppercases the first character of a string.
- Axurldecode: Decodes URL-encoded text using standard query-style decoding.
- Axversion: Returns the current AxonASP runtime version string.
- Axruntimeinfo: Returns a phpinfo-style runtime report with server details, configuration keys, and legal attribution text.
- Axuserhomedirpath: Returns the current user home directory path.
- Axuserconfigdirpath: Returns the resolved full path to config/axonasp.toml.
- Axcachedirpath: Returns the full path to .temp/cache/ with a trailing path separator.
- Axispathseparator: Returns True when the provided single character is a valid path separator for the current platform.
- Axchangetimes: Changes file access and modification timestamps from Unix epoch seconds and returns success as Boolean.
- Axchangemode: Changes file mode using an octal string (for example, 0644) and returns success as Boolean.
- Axcreatelink: Creates a hard link from source to destination and returns success as Boolean.
- Axchangeowner: Changes file owner and group IDs and returns success as Boolean (commonly False on restricted environments).
- Axw: Writes HTML-escaped text to the current response output stream.
- Axwordcount: Counts words in a string or returns the words as an array when format mode is requested.

## Remarks

- Method names are case-insensitive.
- Return type depends on each method operation and arguments.
- For object return values, use Set when assigning the result.
