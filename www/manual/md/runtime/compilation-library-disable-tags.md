# Compile AxonASP with Library Disable Tags

## Overview
This page explains how to compile G3Pix AxonASP with Go build tags that disable optional libraries. Use this when you want a smaller feature surface, fewer dependencies, or stricter runtime exposure. 

Use only the build scripts in the project root for this workflow:
- build.ps1 (Windows)
- build.sh (Linux and macOS)

## Syntax
Windows PowerShell syntax:

```powershell
./build.ps1 -Tags "tag_one tag_two"
```

Linux and macOS Bash syntax:

```bash
./build.sh --tags "tag_one tag_two"
```

Both scripts accept tags separated by spaces, commas, or semicolons.

## Parameters and Arguments
`-Tags` (PowerShell) / `--tags` (Bash):
- Type: String
- Required: No
- Purpose: Passes Go compilation tags to all AxonASP binaries produced by the script.
- Value format: One or more disable tags in a single string.

Supported disable tags:
- lib_adodb_disabled
- lib_adodb_stream_disabled
- lib_g3axonfunctions_disabled
- lib_g3crypto_disabled
- lib_g3db_disabled
- lib_g3fc_disabled
- lib_g3files_disabled
- lib_g3fileuploader_disabled
- lib_g3http_disabled
- lib_g3image_disabled
- lib_g3json_disabled
- lib_g3mail_disabled
- lib_g3md_disabled
- lib_g3pdf_disabled
- lib_g3search_disabled
- lib_g3stringbuilder_disabled
- lib_g3tar_disabled
- lib_g3template_disabled
- lib_g3test_disabled
- lib_g3zip_disabled
- lib_g3zlib_disabled
- lib_g3zstd_disabled
- lib_mswc_disabled
- lib_msxml_disabled
- lib_scripting_dictionary_disabled
- lib_scripting_filesystemobject_disabled
- lib_wscript_shell_disabled

Note about ADOX:
- ADOX is disabled by `lib_adodb_disabled`. This is because ADOX is a sub-library of ADODB and cannot be used if ADODB is disabled. There is no separate `lib_adox_disabled` tag.

## Return Values
The build scripts return:
- Exit code 0 when all selected binaries compile successfully.
- Exit code 1 when at least one target fails to compile.

The script output also shows the active tag string when tags are provided.

## Remarks
- Disable tags are compile-time switches. They do not toggle behavior at runtime.
- If ASP code calls `Server.CreateObject` for a disabled library, object creation will fail with a runtime error.
- Some tags disable foundational COM-compatible objects. Use them only when your application does not depend on those objects.
- For consistent results across all binaries, always build with the same tag set.

## Code Example
Windows examples:

```powershell
# Disable ADODB and MSXML surfaces during compilation
./build.ps1 -Platform windows -Architecture amd64 -Tags "lib_adodb_disabled lib_msxml_disabled"

# Disable image and PDF libraries
./build.ps1 -Tags "lib_g3image_disabled,lib_g3pdf_disabled"
```

Linux and macOS examples:

```bash
# Disable ADODB and MSXML surfaces during compilation
./build.sh --platform linux --arch amd64 --tags "lib_adodb_disabled lib_msxml_disabled"

# Disable image and PDF libraries
./build.sh --tags "lib_g3image_disabled;lib_g3pdf_disabled"
```